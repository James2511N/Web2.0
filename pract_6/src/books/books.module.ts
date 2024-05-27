import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { BooksService } from './books.service';
import { BooksController } from './books.controller';
import { BookSchema } from './schemas/book.schema';

@Module({
  imports: [
    MongooseModule.forFeature([{ name: 'Book', schema: BookSchema }])
  ],
  controllers: [BooksController],
  providers: [BooksService],
})
export class BooksModule {}
