import { z } from "zod";

export const editModeSchema = z.enum(["single", "multi", "global"]);

export const chatTurnSchema = z.object({
  projectId: z.string().min(1),
  deckRevisionId: z.string().min(1),
  target: z.object({
    mode: editModeSchema,
    pageIds: z.array(z.string()).default([]),
  }),
  message: z.string().trim().min(1).max(4000),
  clientRevision: z.number().int().nonnegative().default(0),
  conversationId: z.string().min(1).optional(),
  sessionId: z.string().min(1).optional(),
});

export type ChatTurnInput = z.infer<typeof chatTurnSchema>;

export const createProjectSchema = z.object({
  name: z.string().trim().min(1).max(120),
  themeId: z.string().trim().min(1).max(80).default("fastppt-editorial"),
  themeVersion: z.string().trim().min(1).max(40).default("1.0.0"),
  slidesMarkdown: z.string().max(2_000_000).optional(),
});

export const loginSchema = z.object({
  email: z.string().email().max(240),
  name: z.string().trim().min(1).max(80).optional(),
  accessCode: z.string().max(256).optional(),
});

export const confirmSchema = z.object({
  operationId: z.string().min(1),
});

const editChangeSchema = z.object({
  kind: z.enum([
    "preserve_fact",
    "rewrite_text",
    "layout_change",
    "style_change",
    "image_replace",
    "unsupported",
  ]),
  target: z.string().min(1).max(120),
  value: z.string().max(1000).optional(),
  constraint: z.string().max(500).optional(),
  factId: z.string().max(120).optional(),
});

export const editPlanSchema = z.object({
  workflowMode: z.enum(["import_document", "page_entry", "pptx_beautify"]),
  targetScope: editModeSchema,
  intent: z.string().min(1).max(120),
  affectedPageIds: z.array(z.string().min(1)).max(100),
  pageDelta: z.object({
    add: z.array(z.string().min(1)).max(100),
    remove: z.array(z.string().min(1)).max(100),
    split: z.array(z.string().min(1)).max(100),
    merge: z.array(z.string().min(1)).max(100),
  }),
  changes: z.array(editChangeSchema).max(100),
  factImpact: z.object({
    added: z.array(z.string().max(500)).max(100),
    removed: z.array(z.string().max(500)).max(100),
    changed: z.array(z.string().max(500)).max(100),
  }),
  sourceDocumentIds: z.array(z.string().min(1)).max(100),
  conflictIds: z.array(z.string().min(1)).max(100),
  unsupported: z.array(z.string().max(1000)).max(20),
  requiresConfirmation: z.boolean(),
  confirmationReasons: z.array(z.string().max(1000)).max(20),
  estimatedCost: z.object({
    imageUnits: z.number().nonnegative().max(1000),
    amount: z.number().nonnegative().max(1_000_000),
    currency: z.string().length(3),
  }),
  summary: z.string().min(1).max(2000),
  candidateReasons: z.record(z.string().max(1000)).optional(),
});

export type EditPlanInput = z.infer<typeof editPlanSchema>;

export type CreateProjectInput = z.infer<typeof createProjectSchema>;
