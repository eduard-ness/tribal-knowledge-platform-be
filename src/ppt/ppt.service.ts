import { Injectable, NotFoundException } from '@nestjs/common';
import { Response } from 'express';
import * as fs from 'fs';
import * as path from 'path';
import OpenAI from 'openai';
import pptxgen from 'pptxgenjs';
import { PrismaService } from '../prisma/prisma.service';

type IngestDocumentRecord = {
  title: string | null;
  sourceType: string;
  sourceUrl: string | null;
  content: string;
};

type GeneratePptInput = {
  project: string;
  audience: string;
};

type SlidePlan = {
  title: string;
  subtitle: string;
  slides: {
    title: string;
    bullets: string[];
    speakerNotes?: string;
  }[];
};

@Injectable()
export class PptService {
  private readonly openai: OpenAI;

  constructor(private readonly prisma: PrismaService) {
    if (!process.env.OPENAI_API_KEY) {
      throw new Error('OPENAI_API_KEY is missing');
    }

    this.openai = new OpenAI({
      apiKey: process.env.OPENAI_API_KEY,
    });
  }

  async generatePpt(input: GeneratePptInput) {
    const job = await this.prisma.pptGenerationJob.create({
      data: {
        project: input.project,
        audience: input.audience,
        status: 'processing',
      },
    });

    const documents = await this.prisma.$queryRaw<IngestDocumentRecord[]>`
  SELECT "title", "sourceType", "sourceUrl", "content"
  FROM "IngestDocument"
  ORDER BY "createdAt" DESC
  LIMIT 20
`;

    const content: string = documents
      .map((doc) => {
        const parts: string[] = [
          `Title: ${doc.title ?? 'Untitled'}`,
          `Source type: ${doc.sourceType}`,
          doc.sourceUrl ? `Source URL: ${doc.sourceUrl}` : '',
          doc.content,
        ].filter((value): value is string => Boolean(value));

        return parts.join('\n');
      })
      .join('\n\n---\n\n')
      .slice(0, 45000);

    const slidePlan = await this.generateSlidePlan({
      project: input.project,
      audience: input.audience,
      content,
    });

    const fileName = `${slugify(input.project)}_${slugify(input.audience)}.pptx`;
    const outputDir = this.getOutputDir();
    const outputPath = path.join(outputDir, `${job.id}.pptx`);

    await fs.promises.mkdir(outputDir, { recursive: true });

    await this.renderPowerPoint({
      filePath: outputPath,
      deck: slidePlan,
    });

    let sharePointUrl: string | undefined;

    if (process.env.ENABLE_GRAPH_UPLOAD === 'true') {
      try {
        sharePointUrl = await this.uploadToSharePoint({
          filePath: outputPath,
          fileName,
        });
      } catch (error) {
        console.error(
          'Graph upload failed. Continuing with API download only.',
          error,
        );
      }
    }

    await this.prisma.pptGenerationJob.update({
      where: { id: job.id },
      data: {
        status: 'done',
        fileName,
        sharePointUrl,
      },
    });

    return {
      id: job.id,
      fileName,
      status: 'done',
      downloadUrl: `/ppt/${job.id}/download`,
      sharePointUrl,
    };
  }

  async downloadPpt(id: string, response: Response): Promise<void> {
    const job = await this.prisma.pptGenerationJob.findUnique({
      where: { id },
    });

    if (!job || !job.fileName) {
      throw new NotFoundException(`PowerPoint with id "${id}" was not found`);
    }

    const filePath = path.join(this.getOutputDir(), `${id}.pptx`);

    if (!fs.existsSync(filePath)) {
      throw new NotFoundException(
        `PowerPoint file for id "${id}" was not found`,
      );
    }

    response.download(filePath, job.fileName);
  }

  private async generateSlidePlan(input: {
    project: string;
    audience: string;
    content: string;
  }): Promise<SlidePlan> {
    const fallbackContent =
      input.content ||
      'No ingested content was found. Create a general leadership update about the Tribal Knowledge Platform.';

    const response = await this.openai.responses.create({
      model: 'gpt-4.1-mini',
      input: [
        {
          role: 'system',
          content:
            'You create concise executive PowerPoint outlines. Return only valid JSON matching the requested schema.',
        },
        {
          role: 'user',
          content: `
Create a 5 to 10 slide leadership-ready PowerPoint outline.

Audience: ${input.audience}
Project/topic: ${input.project}

Use this ingested content:
${fallbackContent}

Rules:
- Generate between 5 and 10 slides.
- Include an executive title slide.
- Use concise, leadership-friendly language.
- Each content slide should have 3 to 5 bullets.
- Include speaker notes for each slide.
- Do not invent precise numbers unless they appear in the content.
`,
        },
      ],
      text: {
        format: {
          type: 'json_schema',
          name: 'slide_deck',
          strict: true,
          schema: {
            type: 'object',
            additionalProperties: false,
            required: ['title', 'subtitle', 'slides'],
            properties: {
              title: { type: 'string' },
              subtitle: { type: 'string' },
              slides: {
                type: 'array',
                minItems: 5,
                maxItems: 10,
                items: {
                  type: 'object',
                  additionalProperties: false,
                  required: ['title', 'bullets', 'speakerNotes'],
                  properties: {
                    title: { type: 'string' },
                    bullets: {
                      type: 'array',
                      minItems: 0,
                      maxItems: 5,
                      items: { type: 'string' },
                    },
                    speakerNotes: { type: 'string' },
                  },
                },
              },
            },
          },
        },
      },
    });

    const raw = response.output_text;
    return JSON.parse(raw) as SlidePlan;
  }

  private async renderPowerPoint(input: {
    filePath: string;
    deck: SlidePlan;
  }): Promise<void> {
    const pptx = new pptxgen();

    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = 'Tribal Knowledge Platform';
    pptx.subject = input.deck.title;
    pptx.title = input.deck.title;
    pptx.company = 'Tribal Knowledge Platform';

    this.addTitleSlide(pptx, input.deck.title, input.deck.subtitle);

    input.deck.slides.slice(0, 9).forEach((slide, index) => {
      if (index === 0 && slide.bullets.length === 0) {
        return;
      }

      this.addContentSlide(
        pptx,
        slide.title,
        slide.bullets,
        slide.speakerNotes,
      );
    });

    await pptx.writeFile({ fileName: input.filePath });
  }

  private addTitleSlide(pptx: pptxgen, title: string, subtitle: string): void {
    const slide = pptx.addSlide();

    slide.background = { color: '2F1B59' };

    slide.addText(title, {
      x: 0.7,
      y: 1.45,
      w: 11.8,
      h: 0.8,
      fontSize: 34,
      bold: true,
      color: 'FFFFFF',
      fit: 'shrink',
    });

    slide.addText(subtitle, {
      x: 0.7,
      y: 2.35,
      w: 11.8,
      h: 0.6,
      fontSize: 19,
      color: 'D8C9FF',
      fit: 'shrink',
    });

    slide.addText('Generated from ingested knowledge', {
      x: 0.7,
      y: 6.45,
      w: 11,
      h: 0.35,
      fontSize: 12,
      color: 'FFFFFF',
    });
  }

  private addContentSlide(
    pptx: pptxgen,
    title: string,
    bullets: string[],
    speakerNotes?: string,
  ): void {
    const slide = pptx.addSlide();

    slide.addText(title, {
      x: 0.55,
      y: 0.35,
      w: 12,
      h: 0.55,
      fontSize: 26,
      bold: true,
      color: '2F1B59',
      fit: 'shrink',
    });

    slide.addShape(pptx.ShapeType.line, {
      x: 0.55,
      y: 1.05,
      w: 12.1,
      h: 0,
      line: { color: 'D8C9FF', width: 1 },
    });

    const safeBullets =
      bullets.length > 0
        ? bullets.slice(0, 5)
        : ['No detailed bullet points were generated.'];

    slide.addText(
      safeBullets.map((bullet) => ({
        text: bullet,
        options: { bullet: { indent: 18 } },
      })),
      {
        x: 0.85,
        y: 1.45,
        w: 11.2,
        h: 4.6,
        fontSize: 16,
        color: '323130',
        breakLine: false,
        fit: 'shrink',
      },
    );

    if (speakerNotes) {
      slide.addNotes(speakerNotes);
    }
  }

  private async uploadToSharePoint(input: {
    filePath: string;
    fileName: string;
  }): Promise<string> {
    const tenantId = process.env.GRAPH_TENANT_ID;
    const clientId = process.env.GRAPH_CLIENT_ID;
    const clientSecret = process.env.GRAPH_CLIENT_SECRET;
    const driveId = process.env.GRAPH_DRIVE_ID;
    const folderPath =
      process.env.GRAPH_FOLDER_PATH || 'TribalKnowledgeGenerated';

    if (!tenantId || !clientId || !clientSecret || !driveId) {
      throw new Error('Graph upload env vars are missing');
    }

    const { ClientSecretCredential } = await import('@azure/identity');

    const credential = new ClientSecretCredential(
      tenantId,
      clientId,
      clientSecret,
    );
    const token = await credential.getToken(
      'https://graph.microsoft.com/.default',
    );

    if (!token?.token) {
      throw new Error('Could not acquire Microsoft Graph token');
    }

    const buffer = await fs.promises.readFile(input.filePath);
    const encodedFileName = encodeURIComponent(input.fileName);
    const encodedFolderPath = folderPath
      .split('/')
      .filter(Boolean)
      .map((part) => encodeURIComponent(part))
      .join('/');

    const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedFolderPath}/${encodedFileName}:/content`;

    const response = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token.token}`,
        'Content-Type':
          'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      },
      body: buffer,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Graph upload failed: ${response.status} ${errorText}`);
    }

    const driveItem = (await response.json()) as { webUrl?: string };

    return driveItem.webUrl ?? '';
  }

  private getOutputDir(): string {
    return path.join(process.cwd(), 'generated-ppts');
  }
}

function slugify(value: string): string {
  return value.replace(/[^A-Za-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}
