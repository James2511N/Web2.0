import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { Link } from './link.model';
import { Schema } from 'mongoose';

const LinkSchema = new Schema({
  url: String,
  userId: Number,
});

@Module({
  imports: [MongooseModule.forFeature([{ name: 'Link', schema: LinkSchema }])],
  exports: [MongooseModule],
})
export class LinkModule {}
