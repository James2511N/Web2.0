"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.AppModule = void 0;
const common_1 = require("@nestjs/common");
const users_controller_1 = require("./controllers/users.controller");
const user_service_1 = require("./service/user.service");
const mongoose_1 = require("@nestjs/mongoose");
const schema_1 = require("./schema");
const userAuthorization_middleware_1 = require("./midellware/userAuthorization.middleware");
const orders_controller_1 = require("./controllers/orders.controller");
const service_1 = require("./service");
const adresses_schema_1 = require("./schema/adresses.schema");
let AppModule = class AppModule {
    configure(consumer) {
        consumer.apply(userAuthorization_middleware_1.UserAuthorizationMiddleware).forRoutes('/orders');
    }
};
exports.AppModule = AppModule;
exports.AppModule = AppModule = __decorate([
    (0, common_1.Module)({
        imports: [
            mongoose_1.MongooseModule.forRoot('mongodb+srv://Asstro699:12345qwerty@cluster0.lnr9qbv.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0', { dbName: '4CS-11' }),
            mongoose_1.MongooseModule.forFeature([
                {
                    name: schema_1.Users.name,
                    schema: schema_1.UserSchema,
                },
                {
                    name: schema_1.Orders.name,
                    schema: schema_1.OrdersSchema,
                },
                {
                    name: adresses_schema_1.Adresses.name,
                    schema: adresses_schema_1.AdressesSchema,
                },
            ]),
        ],
        controllers: [users_controller_1.UsersController, orders_controller_1.OrdersController],
        providers: [user_service_1.UserService, service_1.OrderService],
    })
], AppModule);
//# sourceMappingURL=app.module.js.map