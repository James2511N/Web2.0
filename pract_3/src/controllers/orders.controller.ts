import { Body, Controller, Get, Post, Req } from "@nestjs/common";
import { OrderService } from "../service";
import { OrderDto } from "../models";
import { UserLeanDoc } from "../schema";

@Controller({ path: "/orders" })
export class OrdersController {
  constructor(private readonly orderService: OrderService) {}

  @Post("/")
  async createOrder(
    @Body() body: OrderDto,
    @Req() req: Request & { user: UserLeanDoc }
  ) {
    try {
      const { user } = req;
      const order = await this.orderService.createOrder({
        ...body,
        login: user.login,
      });
      return order;
    } catch (err) {
      throw  err;
    }
  }
  @Get("/lowest")
  async getLowestPrice(@Req() req: Request & { user: UserLeanDoc }) {
    try {
      const { user } = req;
      const order = await this.orderService.getLowestPrice(user);
      return order;
    } catch (err) {
      throw err;
    }
  }
  @Get("/biggest")
  async getBiggestPrice(@Req() req: Request & { user: UserLeanDoc }) {
    try {
      const { user } = req;
      const order = await this.orderService.getBiggestPrice(user);
      return order;
    } catch (err) {
      throw err;
    }
  }
  @Get("/last5from")
  async getFiveLastFrom(@Req() req: Request & { user: UserLeanDoc }) {
    try {
      const { user } = req;
      const order = await this.orderService.getFiveLastFrom(user);
      return order;



    } catch (err) {
      throw err;
    }
  }
  @Get("/last3to")
  async getThreeLastTo(@Req() req: Request & { user: UserLeanDoc }) {
    try {
      const { user } = req;
      const order = await this.orderService.getThreeLastTo(user);
      return order;



    } catch (err) {
      throw err;
    }
  }

}
