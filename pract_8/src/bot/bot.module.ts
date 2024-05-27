import { Module } from '@nestjs/common';
import { BotController } from './bot.controller';
import { BotService } from './bot.service';
import { LinkModule } from '../link/link.module';

@Module({
  imports: [LinkModule],
  controllers: [BotController],
  providers: [BotService],
})
export class BotModule {}
