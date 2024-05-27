import { Controller, Get, Post, Body, All, Req, Res } from '@nestjs/common';
import { LoadBalancerService } from './loadBalancer.service';
import { CreateServerDto } from 'src/dto/create-server.dto';
import { Server } from 'src/schemas/server.schema';

@Controller()
export class LoadBalancerController {
  constructor(private readonly loadBalancerService: LoadBalancerService) {}

  @Post('/instances')
  async registerInstance(@Body() createServerDto: CreateServerDto): Promise<void> {
    await this.loadBalancerService.registerInstance(createServerDto);
  }

  @Get('/instances')
  async getInstances(): Promise<Server[]> {
    return await this.loadBalancerService.getInstances();
  }

  @All('/*')
  async redirectRequest(@Req() req, @Res() res) {
    const server = await this.loadBalancerService.getActiveServer();
    if (!server) {
      return res.status(503).send('No active servers available');
    }
    const { method, originalUrl, body, headers } = req;
    const url = `${server.url}${originalUrl}`;
    const response = await this.loadBalancerService.forwardRequest({ url, method, headers, body });
    res.status(response.status).send(response.data);
  }
}
