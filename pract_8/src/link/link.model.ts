import { prop, DocumentType } from '@typegoose/typegoose';

export class Link extends DocumentType {
  @prop({ required: true })
  url: string;

  @prop({ required: true })
  userId: number;
}
