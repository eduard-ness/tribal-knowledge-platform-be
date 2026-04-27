import { Body, Controller, Get, Param, Post, Res } from '@nestjs/common';
import type { Response } from 'express';
import { PptService } from './ppt.service';

@Controller('ppt')
export class PptController {
  constructor(private readonly pptService: PptService) {}

  @Post('generate')
  generatePpt(@Body() body: { project?: string; audience?: string }) {
    return this.pptService.generatePpt({
      project: body.project || 'Tribal Knowledge Platform',
      audience: body.audience || 'Leadership Review',
    });
  }

  @Get(':id/download')
  downloadPpt(@Param('id') id: string, @Res() response: Response) {
    return this.pptService.downloadPpt(id, response);
  }
}
