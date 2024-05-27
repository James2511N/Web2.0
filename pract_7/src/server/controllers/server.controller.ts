import { Controller, Get, Post } from '@nestjs/common';
import { ServerService } from '../services/server.service';
import * as os from 'os';

@Controller()
export class AppController {
  constructor(private readonly serverService: ServerService) {}

  @Post('/orders')
  createOrder(): Promise<void> {
    return this.serverService.createOrder();
  }

  @Get('/api/metrics')
  getMetrics(): { cpu: number, ram: number } {
    const cpuUsage = os.loadavg()[0]; 
    const ramUsage = (os.totalmem() - os.freemem()) / os.totalmem();
    return { cpu: cpuUsage, ram: ramUsage };
  }

  @Get('/api/health')
  getHealth(): string {
    return 'OK';
  }
}
