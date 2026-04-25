import { Module } from '@nestjs/common';
import { PrismaModule } from '../../prisma/prisma.module';
import { IngestController } from './ingest.controller';
import { IngestService } from './ingest.service';

@Module({
  imports: [PrismaModule],
  controllers: [IngestController],
  providers: [IngestService],
})
export class IngestModule {}
