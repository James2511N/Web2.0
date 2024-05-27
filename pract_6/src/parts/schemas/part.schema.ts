import { Schema, Document } from 'mongoose';

export interface Part extends Document {
  imageUrl: string;
  otp: string;
  box: { x: number, y: number, width: number, height: number };
}

export const PartSchema = new Schema({
  imageUrl: { type: String, required: true },
  otp: { type: String, required: true, unique: true },
  box: {
    x: { type: Number, required: true },
    y: { type: Number, required: true },
    width: { type: Number, required: true },
    height: { type: Number, required: true },
  },
});
