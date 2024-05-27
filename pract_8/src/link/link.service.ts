import { Injectable } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { Link } from './link.model';

@Injectable()
export class LinkService {
  constructor(@InjectModel('Link') private readonly linkModel: Model<Link>) {}
}
