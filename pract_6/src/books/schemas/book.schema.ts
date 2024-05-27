import { Schema, Document } from 'mongoose';

export interface Book extends Document {
  title: string;
  pageLinks: { pageLink: string }[];
  userId: string;
}

export const BookSchema = new Schema({
  title: { type: String, required: true },
  pageLinks: [{ pageLink: String }],
  userId: { type: String, required: true },
});
