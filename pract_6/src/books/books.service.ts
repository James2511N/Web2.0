import { Injectable } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { Book } from './schemas/book.schema';
import { CreateBookDto } from './dto/create-book.dto';

@Injectable()
export class BooksService {
  constructor(
    @InjectModel('Book') private readonly bookModel: Model<Book>
  ) {}

  async createBook(createBookDto: CreateBookDto, userId: string): Promise<Book> {
    const newBook = new this.bookModel({ ...createBookDto, userId });
    return newBook.save();
  }

  async getUserBooks(userId: string): Promise<Book[]> {
    return this.bookModel.find({ userId }).exec();
  }
}
