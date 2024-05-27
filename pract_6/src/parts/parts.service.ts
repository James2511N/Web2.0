import { Injectable, BadRequestException } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { Part } from './schemas/part.schema';
import { SubmitPartDto } from './dto/submit-part.dto';
import { v4 as uuidv4 } from 'uuid';

@Injectable()
export class PartsService {
  constructor(
    @InjectModel('Part') private readonly partModel: Model<Part>
  ) {}

  async getRandomPart(): Promise<{ imageUrl: string, otp: string, box: { x: number, y: number, width: number, height: number } }> {
    const parts = await this.partModel.find().exec();
    const part = parts[Math.floor(Math.random() * parts.length)];
    const otp = uuidv4();
    part.otp = otp;
    await part.save();
    return { imageUrl: part.imageUrl, otp, box: part.box };
  }

  async submitPart(partId: string, submitPartDto: SubmitPartDto): Promise<void> {
    const { text, otp } = submitPartDto;
    const part = await this.partModel.findById(partId).exec();
    if (!part || part.otp !== otp) {
      throw new BadRequestException('Invalid OTP or Part ID');
    }
    part.otp = null;
    await part.save();
  }
}
