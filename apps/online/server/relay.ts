import crypto from 'node:crypto';
import { lookup as dnsLookup } from 'node:dns/promises';
import https from 'node:https';
import { isIP } from 'node:net';
import { Readable } from 'node:stream';

const ALLOWED_IMAGE_TYPES = new Set(['image/png', 'image/jpeg', 'image/webp']);
const REDIRECT_STATUSES = new Set([301, 302, 303, 307, 308]);
const DEFAULT_MAX_IMAGE_BYTES = 20 * 1024 * 1024;
const BLOCKED_HOSTS = new Set([
  'localhost',
  'metadata.google.internal',
  'metadata.azure.internal',
  'instance-data.ec2.internal',
]);

export interface RemoteImageDownloadOptions {
  fetchImpl?: typeof fetch;
  lookup?: (hostname: string) => Promise<Array<{ address: string }>>;
  maxBytes?: number;
  maxRedirects?: number;
}

interface RemoteImageResponse {
  status: number;
  ok: boolean;
  headers: Headers;
  body: ReadableStream<Uint8Array> | null;
}

function normalizedMimeType(value: string | null | undefined): string {
  return (value || '').split(';', 1)[0].trim().toLowerCase();
}

function detectedImageType(bytes: Buffer): string | null {
  if (bytes.length >= 8 && bytes.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))) return 'image/png';
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'image/jpeg';
  if (bytes.length >= 12 && bytes.subarray(0, 4).toString('ascii') === 'RIFF' && bytes.subarray(8, 12).toString('ascii') === 'WEBP') return 'image/webp';
  return null;
}

function isBlockedIpv4(address: string): boolean {
  const parts = address.split('.').map(Number);
  if (parts.length !== 4 || parts.some((part) => !Number.isInteger(part) || part < 0 || part > 255)) return true;
  const numeric = parts.reduce((value, part) => (value * 256 + part) >>> 0, 0);
  const inCidr = (base: number[], prefix: number): boolean => {
    const baseNumeric = base.reduce((value, part) => (value * 256 + part) >>> 0, 0);
    const mask = prefix === 0 ? 0 : (0xffffffff << (32 - prefix)) >>> 0;
    return (numeric & mask) === (baseNumeric & mask);
  };
  return inCidr([0, 0, 0, 0], 8)
    || inCidr([10, 0, 0, 0], 8)
    || inCidr([100, 64, 0, 0], 10)
    || inCidr([127, 0, 0, 0], 8)
    || inCidr([169, 254, 0, 0], 16)
    || inCidr([172, 16, 0, 0], 12)
    || inCidr([192, 0, 0, 0], 24)
    || inCidr([192, 0, 2, 0], 24)
    || inCidr([192, 168, 0, 0], 16)
    || inCidr([198, 18, 0, 0], 15)
    || inCidr([198, 51, 100, 0], 24)
    || inCidr([203, 0, 113, 0], 24)
    || inCidr([224, 0, 0, 0], 4);
}

function isBlockedIp(address: string): boolean {
  const normalized = address.toLowerCase().split('%', 1)[0];
  if (isIP(normalized) === 4) return isBlockedIpv4(normalized);
  if (isIP(normalized) !== 6) return true;
  if (normalized.startsWith('::ffff:')) {
    const mapped = normalized.slice('::ffff:'.length);
    if (isIP(mapped) === 4) return isBlockedIpv4(mapped);
    const groups = mapped.split(':');
    if (groups.length === 2) {
      const high = Number.parseInt(groups[0], 16);
      const low = Number.parseInt(groups[1], 16);
      if (Number.isFinite(high) && Number.isFinite(low)) {
        return isBlockedIpv4(`${high >> 8}.${high & 255}.${low >> 8}.${low & 255}`);
      }
    }
    return true;
  }
  return normalized === '::'
    || normalized === '::1'
    || normalized.startsWith('fc')
    || normalized.startsWith('fd')
    || /^fe[89ab]/.test(normalized)
    || normalized.startsWith('ff')
    || normalized.startsWith('2001:db8:');
}

async function validateRemoteImageUrl(url: URL, lookup: NonNullable<RemoteImageDownloadOptions['lookup']>): Promise<Array<{ address: string }>> {
  if (url.protocol !== 'https:') throw new Error('Relay image asset URL must use HTTPS.');
  if (url.username || url.password) throw new Error('Relay image asset URL must not contain credentials.');
  const hostname = url.hostname.toLowerCase().replace(/\.$/, '');
  if (!hostname || BLOCKED_HOSTS.has(hostname) || hostname.endsWith('.local')) throw new Error('Relay image asset host is not allowed.');
  const literal = hostname.startsWith('[') && hostname.endsWith(']') ? hostname.slice(1, -1) : hostname;
  const addresses = isIP(literal) ? [{ address: literal }] : await lookup(hostname);
  if (!addresses.length || addresses.some(({ address }) => isBlockedIp(address))) throw new Error('Relay image asset resolved to a private or non-routable address.');
  return addresses;
}

async function pinnedHttpsRequest(url: URL, addresses: Array<{ address: string }>, signal?: AbortSignal): Promise<RemoteImageResponse> {
  const selected = addresses[0].address;
  const family = isIP(selected);
  return new Promise<RemoteImageResponse>((resolve, reject) => {
    const request = https.request(url, {
      method: 'GET',
      headers: { Accept: 'image/png,image/jpeg,image/webp' },
      lookup: (_hostname, _options, callback) => callback(null, selected, family),
      servername: url.hostname,
      signal,
    }, (response) => {
      const headers = new Headers();
      for (const [name, value] of Object.entries(response.headers)) {
        if (Array.isArray(value)) value.forEach((item) => headers.append(name, item));
        else if (value !== undefined) headers.set(name, String(value));
      }
      resolve({
        status: response.statusCode || 0,
        ok: Boolean(response.statusCode && response.statusCode >= 200 && response.statusCode < 300),
        headers,
        body: Readable.toWeb(response) as ReadableStream<Uint8Array>,
      });
    });
    request.on('error', reject);
    request.end();
  });
}

export async function downloadRemoteImage(urlValue: string, signal?: AbortSignal, options: RemoteImageDownloadOptions = {}): Promise<{ bytes: Buffer; mimeType: string; finalUrl: string }> {
  const lookup = options.lookup || (async (hostname: string) => dnsLookup(hostname, { all: true, verbatim: true }));
  const maxBytes = options.maxBytes || DEFAULT_MAX_IMAGE_BYTES;
  const maxRedirects = options.maxRedirects ?? 3;
  let currentUrl = new URL(urlValue);
  for (let redirectCount = 0; redirectCount <= maxRedirects; redirectCount += 1) {
    const addresses = await validateRemoteImageUrl(currentUrl, lookup);
    const response = options.fetchImpl
      ? await options.fetchImpl(currentUrl, { redirect: 'manual', signal })
      : await pinnedHttpsRequest(currentUrl, addresses, signal);
    if (REDIRECT_STATUSES.has(response.status)) {
      await response.body?.cancel();
      if (redirectCount === maxRedirects) throw new Error(`Relay image asset exceeded ${maxRedirects} redirects.`);
      const location = response.headers.get('location');
      if (!location) throw new Error('Relay image asset redirect did not include a Location header.');
      currentUrl = new URL(location, currentUrl);
      continue;
    }
    if (!response.ok) throw new Error(`Relay image asset returned HTTP ${response.status}.`);
    const mimeType = normalizedMimeType(response.headers.get('content-type'));
    if (!ALLOWED_IMAGE_TYPES.has(mimeType)) {
      await response.body?.cancel();
      throw new Error(`Relay image asset returned unsupported media type: ${mimeType || 'missing'}.`);
    }
    const declaredLength = Number(response.headers.get('content-length'));
    if (Number.isFinite(declaredLength) && declaredLength > maxBytes) {
      await response.body?.cancel();
      throw new Error(`Relay image asset exceeded the ${maxBytes} byte limit.`);
    }
    if (!response.body) throw new Error('Relay image asset returned an empty body.');
    const reader = response.body.getReader();
    const chunks: Buffer[] = [];
    let total = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > maxBytes) {
        await reader.cancel();
        throw new Error(`Relay image asset exceeded the ${maxBytes} byte limit while streaming.`);
      }
      chunks.push(Buffer.from(value));
    }
    if (total === 0) throw new Error('Relay image asset returned an empty body.');
    const bytes = Buffer.concat(chunks, total);
    if (detectedImageType(bytes) !== mimeType) throw new Error('Relay image asset bytes do not match the declared media type.');
    return { bytes, mimeType, finalUrl: currentUrl.toString() };
  }
  throw new Error('Relay image asset redirect validation failed.');
}

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
          const bytes = Buffer.from(item.b64_json, 'base64');
          const mimeType = normalizedMimeType(item.mime_type || 'image/png');
          if (!ALLOWED_IMAGE_TYPES.has(mimeType)) throw new Error(`Relay image response returned unsupported media type: ${mimeType}.`);
          if (bytes.length === 0 || bytes.length > DEFAULT_MAX_IMAGE_BYTES) throw new Error('Relay image response exceeded the 20 MB artifact limit.');
          if (detectedImageType(bytes) !== mimeType) throw new Error('Relay image response bytes do not match the declared media type.');
          return { requestId, model: this.imageModel, bytes, mimeType, width: request.width || 1024, height: request.height || 576 };
        }
        if (!item.url) throw new Error('Relay image response did not include b64_json or url.');
        const image = await downloadRemoteImage(item.url, controller.signal);
        return { requestId, model: this.imageModel, bytes: image.bytes, mimeType: image.mimeType, width: request.width || 1024, height: request.height || 576, sourceUrl: image.finalUrl };
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
