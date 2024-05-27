import { Injectable } from '@nestjs/common';
import { HttpService } from '@nestjs/axios';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { Server } from 'src/schemas/server.schema';
import { CreateServerDto } from 'src/dto/create-server.dto';

@Injectable()
export class LoadBalancerService {
  constructor(
    @InjectModel(Server.name) private serverModel: Model<Server>,
    private httpService: HttpService
  ) {
    this.startMonitoring();
  }

  async registerInstance(createServerDto: CreateServerDto): Promise<void> {
    const newServer = new this.serverModel(createServerDto);
    await newServer.save();
  }

  async getInstances(): Promise<Server[]> {
    return await this.serverModel.find({ status: 'active' }).exec();
  }

  private async checkHealth(server: Server) {
    try {
      const response = await this.httpService.get(`${server.url}/api/health`).toPromise();
      if (response.status !== 200) {
        throw new Error('Health check failed');
      }
      server.status = 'active';
    } catch (error) {
      server.status = 'inactive';
    } finally {
      await server.save();
    }
  }

  private async collectMetrics(server: Server) {
    try {
      const response = await this.httpService.get(`${server.url}/api/metrics`).toPromise();
      const metrics = response.data;
      server.conditions = [
        { type: 'cpu', value: metrics.cpu },
        { type: 'ram', value: metrics.ram },
      ];
    } catch (error) {
      server.status = 'inactive';
    } finally {
      await server.save();
    }
  }

  private startMonitoring() {
    setInterval(async () => {
      const servers = await this.getInstances();
      for (const server of servers) {
        await this.checkHealth(server);
      }
    }, 30000);

    setInterval(async () => {
      const servers = await this.getInstances();
      for (const server of servers) {
        await this.collectMetrics(server);
      }
    }, 60000);
  }

  async getActiveServer(): Promise<Server> {
    const servers = await this.serverModel.find({ status: 'active' }).exec();
    return servers.sort((a, b) => {
      const aCpu = a.conditions.find(cond => cond.type === 'cpu')?.value || Infinity;
      const aRam = a.conditions.find(cond => cond.type === 'ram')?.value || Infinity;
      const bCpu = b.conditions.find(cond => cond.type === 'cpu')?.value || Infinity;
      const bRam = b.conditions.find(cond => cond.type === 'ram')?.value || Infinity;
      return (aCpu + aRam) - (bCpu + bRam);
    })[0];
  }

  async forwardRequest({ url, method, headers, body }) {
    try {
      const response = await this.httpService.request({
        url,
        method,
        headers,
        data: body,
      }).toPromise();
      return response;
    } catch (error) {
      return { status: error.response.status, data: error.response.data };
    }
  }
}
