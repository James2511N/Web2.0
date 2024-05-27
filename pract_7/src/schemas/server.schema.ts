import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
import { Document } from 'mongoose';

@Schema()
class Condition {
  @Prop({ required: true })
  type: string;

  @Prop({ required: true })
  value: any;
}

@Schema()
export class Server extends Document {
  @Prop({ required: true })
  url: string;

  @Prop({ required: true })
  status: string;

  @Prop({ required: true, type: [Condition] })
  conditions: Condition[];
}

export const ServerSchema = SchemaFactory.createForClass(Server);
