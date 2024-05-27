import { Test, TestingModule } from '@nestjs/testing';
import { LoadBalancerController } from './loadBalancer.controller';
import { LoadBalancerService } from './loadBalancer.service';

describe('AppController', () => {
  let appController: LoadBalancerController;

  beforeEach(async () => {
    const app: TestingModule = await Test.createTestingModule({
      controllers: [LoadBalancerController],
      providers: [LoadBalancerService],
    }).compile();

    appController = app.get<LoadBalancerController>(LoadBalancerController);
  });

  describe('root', () => {
    it('should return "Hello World!"', () => {
      expect(appController.getHello()).toBe('Hello World!');
    });
  });
});
