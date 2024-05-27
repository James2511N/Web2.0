import { MiddlewareConsumer, Module, NestModule } from '@nestjs/common'
import { UsersController } from './controllers/users.controller'
import { UserService } from './service/user.service'
import { MongooseModule } from '@nestjs/mongoose'
import { Orders, OrdersSchema, UserSchema, Users } from './schema'
import { UserAuthorizationMiddleware } from './midellware/userAuthorization.middleware'
import { OrdersController } from './controllers/orders.controller'
import { OrderService } from './service'
import { Adresses,AdressesSchema } from './schema/adresses.schema'

@Module({
    imports: [
        MongooseModule.forRoot(
            'mongodb+srv://Asstro699:12345qwerty@cluster0.lnr9qbv.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0',
            { dbName: '4CS-11' }
        ),
        MongooseModule.forFeature([
            {
                name: Users.name,
                schema: UserSchema,
            },
            {
                name: Orders.name,
                schema: OrdersSchema,
            },
            {
              name: Adresses.name,
              schema: AdressesSchema,
          },
        ]),
    ],
    controllers: [UsersController, OrdersController],
    providers: [UserService, OrderService],
})
export class AppModule implements NestModule {
    configure(consumer: MiddlewareConsumer) {
        consumer.apply(UserAuthorizationMiddleware).forRoutes('/orders')
    }
}
