"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var __param = (this && this.__param) || function (paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.OrdersController = void 0;
const common_1 = require("@nestjs/common");
const service_1 = require("../service");
const models_1 = require("../models");
let OrdersController = class OrdersController {
    constructor(orderService) {
        this.orderService = orderService;
    }
    async createOrder(body, req) {
        try {
            const { user } = req;
            const order = await this.orderService.createOrder({
                ...body,
                login: user.login,
            });
            return order;
        }
        catch (err) {
            throw err;
        }
    }
    async getLowestPrice(req) {
        try {
            const { user } = req;
            const order = await this.orderService.getLowestPrice(user);
            return order;
        }
        catch (err) {
            throw err;
        }
    }
    async getBiggestPrice(req) {
        try {
            const { user } = req;
            const order = await this.orderService.getBiggestPrice(user);
            return order;
        }
        catch (err) {
            throw err;
        }
    }
    async getFiveLastFrom(req) {
        try {
            const { user } = req;
            const order = await this.orderService.getFiveLastFrom(user);
            return order;
        }
        catch (err) {
            throw err;
        }
    }
    async getThreeLastTo(req) {
        try {
            const { user } = req;
            const order = await this.orderService.getThreeLastTo(user);
            return order;
        }
        catch (err) {
            throw err;
        }
    }
};
exports.OrdersController = OrdersController;
__decorate([
    (0, common_1.Post)("/"),
    __param(0, (0, common_1.Body)()),
    __param(1, (0, common_1.Req)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [models_1.OrderDto, Object]),
    __metadata("design:returntype", Promise)
], OrdersController.prototype, "createOrder", null);
__decorate([
    (0, common_1.Get)("/lowest"),
    __param(0, (0, common_1.Req)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", Promise)
], OrdersController.prototype, "getLowestPrice", null);
__decorate([
    (0, common_1.Get)("/biggest"),
    __param(0, (0, common_1.Req)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", Promise)
], OrdersController.prototype, "getBiggestPrice", null);
__decorate([
    (0, common_1.Get)("/last5from"),
    __param(0, (0, common_1.Req)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", Promise)
], OrdersController.prototype, "getFiveLastFrom", null);
__decorate([
    (0, common_1.Get)("/last3to"),
    __param(0, (0, common_1.Req)()),
    __metadata("design:type", Function),
    __metadata("design:paramtypes", [Object]),
    __metadata("design:returntype", Promise)
], OrdersController.prototype, "getThreeLastTo", null);
exports.OrdersController = OrdersController = __decorate([
    (0, common_1.Controller)({ path: "/orders" }),
    __metadata("design:paramtypes", [service_1.OrderService])
], OrdersController);
//# sourceMappingURL=orders.controller.js.map