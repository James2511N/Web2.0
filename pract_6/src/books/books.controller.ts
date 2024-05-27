import { Controller, Post, Get, Body, Request, UseGuards } from '@nestjs/common';
import { BooksService } from './books.service';
import { CreateBookDto } from './dto/create-book.dto';
import { AuthenticatedGuard } from '../auth/authenticated.guard';

@Controller('books')
export class BooksController {
  constructor(private readonly booksService: BooksService) {}

  @UseGuards(AuthenticatedGuard)
  @Post()
  async createBook(@Body() createBookDto: CreateBookDto, @Request() req) {
    return this.booksService.createBook(createBookDto, req.user._id);
  }

  @UseGuards(AuthenticatedGuard)
  @Get()
  async getUserBooks(@Request() req) {
    return this.booksService.getUserBooks(req.user._id);
  }
}
