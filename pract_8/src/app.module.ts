import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { BotModule } from './bot/bot.module';

@Module({
  imports: [
    MongooseModule.forRoot('mongodb+srv://mynameseriy:bw66fv5wwKJx7djz@tg.0jd2kdu.mongodb.net/?retryWrites=true&w=majority&appName=tg'),
    BotModule,
  ],
})
export class AppModule {}
