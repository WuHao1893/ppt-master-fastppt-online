import crypto from 'node:crypto';

export interface RelayStatus {
  configured: boolean;
  baseUrl: string | null;
  model: string;
  requestTimeoutMs: number;
  priceSnapshot: { imageUnit: number; tokenUnit: number };
  imageModel: string;
}

interface StructuredRequest {
  system: string;
  user: string;
  imageUrls?: string[];
}

export interface ImageGenerationResult {
  requestId: string;
  model: string;
  bytes: Buffer;
  mimeType: string;
  width: number;
  height: number;
  sourceUrl?: string;
}

export class RelayModelAdapter {
  private readonly baseUrl = process.env.RELAY_BASE_URL?.trim().replace(/\/$/, '') || '';
  private readonly apiKey = process.env.RELAY_API_KEY?.trim() || '';
  private readonly model = process.env.RELAY_MODEL || 'fastppt-online-planner';
  private readonly imageModel = process.env.RELAY_IMAGE_MODEL || process.env.RELAY_MODEL || 'fastppt-online-image';
  private readonly timeoutMs = Number(process.env.RELAY_TIMEOUT_MS || 12_000);
  private readonly priceSnapshot = {
    imageUnit: Number(process.env.RELAY_PRICE_IMAGE_UNIT || 0.08),
    tokenUnit: Number(process.env.RELAY_PRICE_TOKEN_UNIT || 0),
  };

  status(): RelayStatus {
    return { configured: Boolean(this.baseUrl && this.apiKey), baseUrl: this.baseUrl || null, model: this.model, imageModel: this.imageModel, requestTimeoutMs: this.timeoutMs, priceSnapshot: this.priceSnapshot };
  }

  async structuredJson<T>(request: StructuredRequest): Promise<{ value: T; requestId: string; model: string }> {
    if (!this.baseUrl || !this.apiKey) throw new Error('Relay model adapter is not configured.');
    const requestId = `relay_${crypto.randomUUID().replaceAll('-', '')}`;
    let lastError: Error | null = null;
    for (let attempt = 0; attempt < 2; attempt += 1) {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), this.timeoutMs);
      try {
        const userContent = request.imageUrls?.length
          ? [{ type: 'text', text: request.user }, ...request.imageUrls.map((url) => ({ type: 'image_url', image_url: { url } }))]
          : request.user;
        const response = await fetch(`${this.baseUrl}/v1/chat/completions`, {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${this.apiKey}`,
            'Content-Type': 'application/json',
            'X-Request-ID': requestId,
          },
          body: JSON.stringify({
            model: this.model,
            messages: [{ role: 'system', content: request.system }, { role: 'user', content: userContent }],
            response_format: { type: 'json_object' },
            stream: false,
          }),
          signal: controller.signal,
        });
        if (!response.ok) throw new Error(`Relay returned HTTP ${response.status}.`);
        const envelope = await response.json() as { choices?: Array<{ message?: { content?: string } }> };
        const content = envelope.choices?.[0]?.message?.content;
        if (!content) throw new Error('Relay response did not include structured content.');
        return { value: JSON.parse(content) as T, requestId, model: this.model };
      } catch (error) {
        lastError = error as Error;
        if (attempt === 0) await new Promise((resolve) => setTimeout(resolve, 150));
      } finally {
        clearTimeout(timer);
      }
    }
    throw new Error(`Relay request ${requestId} failed: ${lastError?.message || 'unknown error'}`);
  }

  async generateImage(request: { prompt: string; width?: number; height?: number; quality?: string }): Promise<ImageGenerationResult> {
    if (!this.baseUrl || !this.apiKey) throw new Error('Relay image adapter is not configured. Set RELAY_BASE_URL and RELAY_API_KEY.');
    const requestId = `relay_${crypto.randomUUID().replaceAll('-', '')}`;
    let lastError: Error | null = null;
    for (let attempt = 0; attempt < 2; attempt += 1) {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), this.timeoutMs);
      try {
        const response = await fetch(`${this.baseUrl}/v1/images`, {
          method: 'POST',
          headers: { Authorization: `Bearer ${this.apiKey}`, 'Content-Type': 'application/json', 'X-Request-ID': requestId },
          body: JSON.stringify({
            model: this.imageModel,
            prompt: request.prompt,
            size: `${request.width || 1024}x${request.height || 576}`,
            quality: request.quality || 'standard',
            response_format: 'b64_json',
          }),
          signal: controller.signal,
        });
        if (!response.ok) throw new Error(`Relay image request returned HTTP ${response.status}.`);
        const envelope = await response.json() as { data?: Array<{ b64_json?: string; url?: string; mime_type?: string }> };
        const item = envelope.data?.[0];
        if (!item) throw new Error('Relay image response did not include an image.');
        if (item.b64_json) {
          return { requestId, model: this.imageModel, bytes: Buffer.from(item.b64_json, 'base64'), mimeType: item.mime_type || 'image/png', width: request.width || 1024, height: request.height || 576 };
        }
        if (!item.url) throw new Error('Relay image response did not include b64_json or url.');
        const imageResponse = await fetch(item.url, { signal: controller.signal });
        if (!imageResponse.ok) throw new Error(`Relay image asset returned HTTP ${imageResponse.status}.`);
        const bytes = Buffer.from(await imageResponse.arrayBuffer());
        return { requestId, model: this.imageModel, bytes, mimeType: item.mime_type || imageResponse.headers.get('content-type') || 'image/png', width: request.width || 1024, height: request.height || 576, sourceUrl: item.url };
      } catch (error) {
        lastError = error as Error;
        if (attempt === 0) await new Promise((resolve) => setTimeout(resolve, 150));
      } finally {
        clearTimeout(timer);
      }
    }
    throw new Error(`Relay image request ${requestId} failed: ${lastError?.message || 'unknown error'}`);
  }
}
