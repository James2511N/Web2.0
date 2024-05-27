import { Injectable } from '@nestjs/common';
import { getRandomTime } from '../shared/getRandomTime';

@Injectable()
export class ServerService {
  private minTime = 30;
  private maxTime = 100

  async createOrder(): Promise<void> {
    return new Promise((res) => setTimeout(res, getRandomTime(this.minTime, this.maxTime) * 1000))
  }
}
