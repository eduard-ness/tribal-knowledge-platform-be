import { Module } from '@nestjs/common';
import { PrismaModule } from '../prisma/prisma.module';
import { PptController } from './ppt.controller';
import { PptService } from './ppt.service';

@Module({
  imports: [PrismaModule],
  controllers: [PptController],
  providers: [PptService],
})
export class PptModule {}
