import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { PrismaModule } from './prisma/prisma.module';
import { IngestModule } from './ingest/ingest.module';
import { PptModule } from './ppt/ppt.module';

@Module({
  imports: [PrismaModule, IngestModule, PptModule],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
