import { Injectable, NotFoundException } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';

type UploadedLocalFile = {
  originalname: string;
  mimetype?: string;
  size?: number;
};

type StartIngestionInput = {
  files: UploadedLocalFile[];
  urls: string[];
  sources: string[];
};

@Injectable()
export class IngestService {
  constructor(private readonly prisma: PrismaService) {}

  async startIngestion(input: StartIngestionInput) {
    const selectedSources = new Set(input.sources);

    const uploadedItems = [
      ...this.mapLocalFiles(input.files),
      ...(selectedSources.has('sharepoint')
        ? this.getMockSharePointFiles()
        : []),
      ...this.mapUrls(input.urls),
    ];

    const job = await this.prisma.ingestJob.create({
      data: {
        status: 'queued',
        progress: 0,
        ...(uploadedItems.length > 0
          ? {
              files: {
                create: uploadedItems,
              },
            }
          : {}),
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

  private mapLocalFiles(files: UploadedLocalFile[]) {
    return files.map((file) => ({
      fileName: file.originalname,
      contentType: file.mimetype ?? 'application/octet-stream',
      size: file.size ?? 0,
    }));
  }

  private mapUrls(urls: string[]) {
    return urls
      .map((url) => url.trim())
      .filter(Boolean)
      .map((url) => ({
        fileName: url,
        contentType: 'text/url',
        size: 0,
      }));
  }

  private getMockSharePointFiles() {
    return [
      {
        fileName: 'SharePoint_Project_Overview.docx',
        contentType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        size: 12400,
      },
      {
        fileName: 'SharePoint_Architecture_Notes.pdf',
        contentType: 'application/pdf',
        size: 48200,
      },
      {
        fileName: 'SharePoint_Risk_Register.xlsx',
        contentType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        size: 19800,
      },
    ];
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
