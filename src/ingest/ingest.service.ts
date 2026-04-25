import { Injectable, NotFoundException } from '@nestjs/common';
import { PrismaService } from '../../prisma/prisma.service';

type UploadedFileInput = {
  originalname: string;
  mimetype?: string;
  size?: number;
};

@Injectable()
export class IngestService {
  constructor(private readonly prisma: PrismaService) {}

  async createJob() {
    const job = await this.prisma.ingestJob.create({
      data: {
        status: 'queued',
        progress: 0,
      },
    });

    void this.simulateProgress(job.id);

    return job;
  }

  async createUploadJob(files: UploadedFileInput[], urls: string[] = []) {
    const job = await this.prisma.ingestJob.create({
      data: {
        status: 'queued',
        progress: 0,
        files: {
          create: [
            ...files.map((file) => ({
              fileName: file.originalname,
              contentType: file.mimetype,
              size: file.size,
            })),
            ...urls.map((url) => ({
              fileName: url,
              contentType: 'text/url',
              size: 0,
            })),
          ],
        },
      },
      include: {
        files: true,
      },
    });

    void this.simulateProgress(job.id);

    return job;
  }

  async getJob(id: string) {
    const job = await this.prisma.ingestJob.findUnique({
      where: { id },
      include: {
        files: true,
      },
    });

    if (!job) {
      throw new NotFoundException(`Ingest job with id "${id}" was not found`);
    }

    return job;
  }

  private async simulateProgress(id: string): Promise<void> {
    await this.prisma.ingestJob.update({
      where: { id },
      data: {
        status: 'processing',
        progress: 5,
      },
    });

    let progress = 5;

    const interval = setInterval(() => {
      void (async () => {
        try {
          progress += 10;

          if (progress >= 100) {
            await this.prisma.ingestJob.update({
              where: { id },
              data: {
                status: 'done',
                progress: 100,
              },
            });

            clearInterval(interval);
            return;
          }

          await this.prisma.ingestJob.update({
            where: { id },
            data: {
              status: 'processing',
              progress,
            },
          });
        } catch (error) {
          clearInterval(interval);
          console.error(`Failed to update ingest job "${id}"`, error);
        }
      })();
    }, 1000);
  }
}
