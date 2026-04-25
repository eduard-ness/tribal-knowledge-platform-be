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

@Controller('ingest')
export class IngestController {
  constructor(private readonly ingestService: IngestService) {}

  @Post()
  createJob() {
    return this.ingestService.createJob();
  }

  @Post('upload')
  @UseInterceptors(FilesInterceptor('files'))
  upload(
    @UploadedFiles() files: Express.Multer.File[],
    @Body('urls') urls?: string | string[],
  ) {
    const normalizedUrls = Array.isArray(urls) ? urls : urls ? [urls] : [];

    return this.ingestService.createUploadJob(files ?? [], normalizedUrls);
  }

  @Get(':id')
  getJob(@Param('id') id: string) {
    return this.ingestService.getJob(id);
  }
}
