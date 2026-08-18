import type { DurableJob } from '../shared/models.js';
import type { StateStore } from './store.js';
import { nowIso } from './utils.js';

type Runner = () => Promise<void>;

interface Waiter {
  resolve: () => void;
  reject: (error: unknown) => void;
}

/**
 * A small durable dispatcher for the API process. The job row is written
 * before a runner is scheduled, so a restart can safely re-enqueue queued or
 * interrupted jobs. The concurrency limit prevents one project from
 * exhausting the worker process while keeping the MVP deployable without a
 * second queue service.
 */
export class DurableJobQueue {
  private readonly pending: string[] = [];
  private readonly queued = new Set<string>();
  private readonly runners = new Map<string, Runner>();
  private readonly waiters = new Map<string, Waiter[]>();
  private active = 0;
  private pumping = false;
  private readonly concurrency: number;

  constructor(private readonly store: StateStore, concurrency = Number(process.env.JOB_CONCURRENCY || 2)) {
    this.concurrency = Number.isFinite(concurrency) && concurrency > 0 ? Math.min(Math.floor(concurrency), 16) : 2;
  }

  async enqueue(job: DurableJob, runner: Runner, wait = false): Promise<void> {
    let shouldSchedule = true;
    await this.store.update((state) => {
      const existing = state.jobs.find((candidate) => candidate.jobId === job.jobId);
      if (existing) {
        if (existing.status === 'completed' || existing.status === 'failed') {
          shouldSchedule = false;
          return;
        }
        existing.status = 'queued';
        existing.completedAt = null;
      } else {
        state.jobs.push({ ...job, status: 'queued', completedAt: null });
      }
    });
    if (!shouldSchedule) return;
    this.runners.set(job.jobId, runner);
    if (!this.queued.has(job.jobId)) {
      this.queued.add(job.jobId);
      this.pending.push(job.jobId);
    }
    const completion = wait ? new Promise<void>((resolve, reject) => {
      const list = this.waiters.get(job.jobId) || [];
      list.push({ resolve, reject });
      this.waiters.set(job.jobId, list);
    }) : null;
    void this.pump();
    if (completion) await completion;
  }

  async resume(jobs: DurableJob[], resolveRunner: (job: DurableJob) => Runner | null): Promise<void> {
    for (const job of jobs.filter((candidate) => candidate.status === 'queued' || candidate.status === 'running')) {
      const runner = resolveRunner(job);
      if (!runner) {
        await this.store.update((state) => {
          const current = state.jobs.find((candidate) => candidate.jobId === job.jobId);
          if (current) { current.status = 'failed'; current.completedAt = nowIso(); }
        });
        continue;
      }
      await this.enqueue({ ...job, status: 'queued' }, runner);
    }
  }

  private async pump(): Promise<void> {
    if (this.pumping) return;
    this.pumping = true;
    try {
      while (this.active < this.concurrency && this.pending.length) {
        const jobId = this.pending.shift()!;
        this.queued.delete(jobId);
        const runner = this.runners.get(jobId);
        if (!runner) continue;
        const job = this.store.state.jobs.find((candidate) => candidate.jobId === jobId);
        if (!job || job.status === 'completed' || job.status === 'failed') {
          this.finish(jobId);
          continue;
        }
        this.active += 1;
        await this.store.update((state) => {
          const current = state.jobs.find((candidate) => candidate.jobId === jobId);
          if (current) current.status = 'running';
        });
        void runner().then(
          () => this.finish(jobId),
          (error) => this.finish(jobId, error),
        );
      }
    } finally {
      this.pumping = false;
    }
  }

  private finish(jobId: string, error?: unknown): void {
    this.active = Math.max(0, this.active - 1);
    this.runners.delete(jobId);
    const waiters = this.waiters.get(jobId) || [];
    this.waiters.delete(jobId);
    void this.store.update((state) => {
      const job = state.jobs.find((candidate) => candidate.jobId === jobId);
      if (job && job.status === 'running') {
        job.status = error ? 'failed' : 'completed';
        job.completedAt = nowIso();
      }
    }).then(
      () => waiters.forEach((waiter) => error ? waiter.reject(error) : waiter.resolve()),
      (persistError) => waiters.forEach((waiter) => waiter.reject(persistError)),
    ).finally(() => { void this.pump(); });
  }
}
