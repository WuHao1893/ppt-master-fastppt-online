import { randomUUID } from "node:crypto";
import type { DurableJob } from "../shared/models.js";
import type { StateStore } from "./store.js";

type Runner = () => Promise<void>;
type RunnerResolver = (job: DurableJob) => Runner | null;

interface Waiter {
  resolve: () => void;
  reject: (error: unknown) => void;
}

export class DurableJobQueue {
  private readonly runners = new Map<string, Runner>();
  private readonly waiters = new Map<string, Waiter[]>();
  private readonly active = new Set<string>();
  private readonly workerId = `worker_${process.pid}_${randomUUID()}`;
  private readonly concurrency: number;
  private readonly leaseMs: number;
  private readonly pollMs: number;
  private resolver: RunnerResolver | null = null;
  private pumping = false;
  private timer: NodeJS.Timeout | null = null;

  constructor(
    private readonly store: StateStore,
    concurrency = Number(process.env.JOB_CONCURRENCY || 2),
  ) {
    this.concurrency =
      Number.isFinite(concurrency) && concurrency > 0
        ? Math.min(Math.floor(concurrency), 16)
        : 2;
    this.leaseMs = Math.max(
      5_000,
      Number(process.env.JOB_LEASE_MS || 30_000),
    );
    this.pollMs = Math.max(100, Number(process.env.JOB_POLL_MS || 500));
  }

  async enqueue(job: DurableJob, runner: Runner, wait = false): Promise<void> {
    const durable = await this.store.enqueueJob(job);
    if (durable.status === "completed") return;
    if (durable.status === "failed")
      throw new Error(durable.lastError || `Job ${durable.jobId} already failed.`);
    this.runners.set(durable.jobId, runner);
    const completion = wait
      ? new Promise<void>((resolve, reject) => {
          const list = this.waiters.get(durable.jobId) || [];
          list.push({ resolve, reject });
          this.waiters.set(durable.jobId, list);
        })
      : null;
    this.schedulePump();
    if (completion) await completion;
  }

  async resume(
    _jobs: DurableJob[],
    resolveRunner: RunnerResolver,
  ): Promise<void> {
    this.resolver = resolveRunner;
    await this.store.releaseExpiredJobLeases();
    this.schedulePump();
  }

  private schedulePump(delayMs = 0): void {
    if (this.timer) return;
    this.timer = setTimeout(() => {
      this.timer = null;
      void this.pump();
    }, Math.max(0, delayMs));
    this.timer.unref?.();
  }

  private async pump(): Promise<void> {
    if (this.pumping) return;
    if (!this.resolver && this.runners.size === 0) return;
    this.pumping = true;
    try {
      while (this.active.size < this.concurrency) {
        const job = await this.store.claimJob(this.workerId, this.leaseMs);
        if (!job) break;
        const runner = this.runners.get(job.jobId) || this.resolver?.(job);
        if (!runner) {
          const failed = await this.store.failJob(
            job.jobId,
            this.workerId,
            `No runner is registered for ${job.kind} job ${job.jobId}.`,
            0,
          );
          if (failed?.status === "failed")
            this.settleWaiters(
              job.jobId,
              new Error(failed.lastError || "Job runner is unavailable."),
            );
          continue;
        }
        this.active.add(job.jobId);
        void this.runClaimed(job, runner);
      }
    } finally {
      this.pumping = false;
      if (this.resolver || this.runners.size > 0) this.schedulePump(this.pollMs);
    }
  }

  private async runClaimed(job: DurableJob, runner: Runner): Promise<void> {
    const heartbeat = setInterval(() => {
      void this.store.heartbeatJob(job.jobId, this.workerId, this.leaseMs);
    }, Math.max(1_000, Math.floor(this.leaseMs / 3)));
    heartbeat.unref?.();
    try {
      await runner();
      const completed = await this.store.completeJob(job.jobId, this.workerId);
      if (completed) {
        this.runners.delete(job.jobId);
        this.settleWaiters(job.jobId);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      const retryDelay = Math.min(
        30_000,
        500 * 2 ** Math.max(0, job.attemptCount - 1),
      );
      const failed = await this.store.failJob(
        job.jobId,
        this.workerId,
        message,
        retryDelay,
      );
      if (failed?.status === "failed") {
        this.runners.delete(job.jobId);
        this.settleWaiters(job.jobId, error);
      } else if (failed) {
        const delay = Math.max(
          0,
          new Date(failed.availableAt).getTime() - Date.now(),
        );
        this.schedulePump(delay);
      }
    } finally {
      clearInterval(heartbeat);
      this.active.delete(job.jobId);
      if (this.resolver || this.runners.size > 0) this.schedulePump();
    }
  }

  private settleWaiters(jobId: string, error?: unknown): void {
    const waiters = this.waiters.get(jobId) || [];
    this.waiters.delete(jobId);
    for (const waiter of waiters) {
      if (error) waiter.reject(error);
      else waiter.resolve();
    }
  }
}
