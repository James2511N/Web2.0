import { Test, TestingModule } from '@nestjs/testing';
import { AppController } from './server.controller';
import { ServerService } from '../services/server.service';

describe('AppController', () => {
  let appController: AppController;

  beforeEach(async () => {
    const app: TestingModule = await Test.createTestingModule({
      controllers: [AppController],
      providers: [ServerService],
    }).compile();

    appController = app.get<AppController>(AppController);
  });
});
