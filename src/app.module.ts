import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { PrismaModule } from '../prisma/prisma.module';
import { IngestModule } from './ingest/ingest.module';

@Module({
  imports: [PrismaModule, IngestModule],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
