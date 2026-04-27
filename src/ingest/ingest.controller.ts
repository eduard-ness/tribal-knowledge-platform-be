import {
  Body,
  Controller,
  Get,
  Param,
  Post,
  UploadedFiles,
  UseInterceptors,
} from '@nestjs/common';
import { FilesInterceptor } from '@nestjs/platform-express';
import { IngestService } from './ingest.service';

type UploadedLocalFile = {
  originalname: string;
  mimetype?: string;
  size?: number;
};

function normalizeArray(value?: string | string[]): string[] {
  if (!value) {
    return [];
  }

  return Array.isArray(value) ? value : [value];
}

@Controller('ingest')
export class IngestController {
  constructor(private readonly ingestService: IngestService) {}

  @Post('start')
  @UseInterceptors(FilesInterceptor('files', 50))
  startIngestion(
    @UploadedFiles() files: UploadedLocalFile[] = [],
    @Body('urls') urls?: string | string[],
    @Body('sources') sources?: string | string[],
  ) {
    return this.ingestService.startIngestion({
      files,
      urls: normalizeArray(urls),
      sources: normalizeArray(sources),
    });
  }

  @Get(':id')
  getJob(@Param('id') id: string) {
    return this.ingestService.getJob(id);
  }
}
