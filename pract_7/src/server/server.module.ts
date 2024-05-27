import { Module } from '@nestjs/common';
import { AppController } from './controllers/server.controller';
import { ServerService } from './services/server.service';
import { ConfigModule } from '@nestjs/config';
import configuration from './config/configuration';

@Module({
  imports: [ConfigModule.forRoot({
    load: [configuration],
  })],
  controllers: [AppController],
  providers: [ServerService],
})
export class ServerModule {}
