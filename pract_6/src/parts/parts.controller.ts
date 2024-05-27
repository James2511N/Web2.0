import { Controller, Get, Post, Param, Body, UseGuards } from '@nestjs/common';
import { PartsService } from './parts.service';
import { SubmitPartDto } from './dto/submit-part.dto';
import { AuthenticatedGuard } from '../auth/authenticated.guard';

@Controller('parts')
export class PartsController {
  constructor(private readonly partsService: PartsService) {}

  @UseGuards(AuthenticatedGuard)
  @Get()
  async getRandomPart() {
    return this.partsService.getRandomPart();
  }

  @UseGuards(AuthenticatedGuard)
  @Post(':id')
  async submitPart(@Param('id') partId: string, @Body() submitPartDto: SubmitPartDto) {
    return this.partsService.submitPart(partId, submitPartDto);
  }
}
