import fs from 'node:fs/promises';
import path from 'node:path';
import { GetObjectCommand, PutObjectCommand, S3Client } from '@aws-sdk/client-s3';

export interface StoredArtifact {
  objectKey: string;
  localPath?: string;
}

export interface ObjectStoreStatus {
  kind: 'local' | 's3';
  bucket: string | null;
  endpoint: string | null;
}

export class ObjectStorage {
  private readonly client: S3Client | null;
  private readonly bucket: string;
  private readonly root: string;
  readonly status: ObjectStoreStatus;

  constructor() {
    const endpoint = process.env.OBJECT_STORAGE_ENDPOINT?.trim() || '';
    const bucket = process.env.OBJECT_STORAGE_BUCKET?.trim() || '';
    const accessKeyId = process.env.OBJECT_STORAGE_ACCESS_KEY?.trim() || '';
    const secretAccessKey = process.env.OBJECT_STORAGE_SECRET_KEY?.trim() || '';
    this.bucket = bucket;
    this.root = path.resolve(process.env.DATA_DIR || './data', 'objects');
    if (endpoint && bucket && accessKeyId && secretAccessKey) {
      this.client = new S3Client({
        region: process.env.OBJECT_STORAGE_REGION || 'us-east-1',
        endpoint,
        forcePathStyle: process.env.OBJECT_STORAGE_FORCE_PATH_STYLE !== 'false',
        credentials: { accessKeyId, secretAccessKey },
      });
      this.status = { kind: 's3', bucket, endpoint };
    } else {
      if (process.env.NODE_ENV === 'production') throw new Error('S3-compatible object storage configuration is required in production.');
      this.client = null;
      this.status = { kind: 'local', bucket: null, endpoint: null };
    }
  }

  async putFile(localPath: string, objectKey: string, contentType: string): Promise<StoredArtifact> {
    if (this.client) {
      const body = await fs.readFile(localPath);
      await this.client.send(new PutObjectCommand({ Bucket: this.bucket, Key: objectKey, Body: body, ContentType: contentType }));
      return { objectKey };
    }
    const target = path.resolve(this.root, objectKey);
    if (!target.startsWith(`${this.root}${path.sep}`)) throw new Error('Object key escaped local object storage root.');
    await fs.mkdir(path.dirname(target), { recursive: true });
    await fs.copyFile(localPath, target);
    return { objectKey, localPath: target };
  }

  async putBuffer(bytes: Buffer, objectKey: string, contentType: string): Promise<StoredArtifact> {
    if (this.client) {
      await this.client.send(new PutObjectCommand({ Bucket: this.bucket, Key: objectKey, Body: bytes, ContentType: contentType }));
      return { objectKey };
    }
    const target = path.resolve(this.root, objectKey);
    if (!target.startsWith(`${this.root}${path.sep}`)) throw new Error('Object key escaped local object storage root.');
    await fs.mkdir(path.dirname(target), { recursive: true });
    await fs.writeFile(target, bytes, { flag: 'wx' });
    return { objectKey, localPath: target };
  }

  async getBytes(objectKey: string): Promise<Buffer> {
    if (this.client) {
      const response = await this.client.send(new GetObjectCommand({ Bucket: this.bucket, Key: objectKey }));
      if (!response.Body) throw new Error('Object storage returned an empty body.');
      return Buffer.from(await response.Body.transformToByteArray());
    }
    const target = path.resolve(this.root, objectKey);
    if (!target.startsWith(`${this.root}${path.sep}`)) throw new Error('Object key escaped local object storage root.');
    return fs.readFile(target);
  }
}
