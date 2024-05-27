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
exports.OrderService = void 0;
const common_1 = require("@nestjs/common");
const schema_1 = require("../schema");
const adresses_schema_1 = require("../schema/adresses.schema");
const mongoose_1 = require("mongoose");
const mongoose_2 = require("@nestjs/mongoose");
const shared_1 = require("../shared");
let OrderService = class OrderService {
    constructor(orderModel, addressModel) {
        this.orderModel = orderModel;
        this.addressModel = addressModel;
    }
    async createOrder(body) {
        const fromAddressExists = await this.addressModel.exists({
            name: body.from,
        });
        const toAddressExists = await this.addressModel.exists({
            name: body.to,
        });
        if (!fromAddressExists || !toAddressExists) {
            throw new Error('Either the From or To address is invalid and not found in our records.');
        }
        const fromAddressDocument = await this.addressModel.findOne({ name: body.from });
        const toAddressDocument = await this.addressModel.findOne({ name: body.to });
        function distance() {
            const latitudeFrom = fromAddressDocument
                ? fromAddressDocument.location.latitude
                : null;
            const latitudeTo = toAddressDocument
                ? toAddressDocument.location.latitude
                : null;
            const longitudeFrom = fromAddressDocument
                ? fromAddressDocument.location.longitude
                : null;
            const longitudeTo = toAddressDocument
                ? toAddressDocument.location.longitude
                : null;
            var radlat1 = (Math.PI * latitudeFrom) / 180;
            var radlat2 = (Math.PI * latitudeTo) / 180;
            var theta = longitudeFrom - longitudeTo;
            var radtheta = (Math.PI * theta) / 180;
            var dist = Math.sin(radlat1) * Math.sin(radlat2) +
                Math.cos(radlat1) * Math.cos(radlat2) * Math.cos(radtheta);
            if (dist > 1) {
                dist = 1;
            }
            dist = Math.acos(dist);
            dist = (dist * 180) / Math.PI;
            dist = dist * 60 * 1.1515;
            return Math.round(dist * 1.609344);
        }
        const randomNumber = Math.random();
        const randomInRange = Math.floor(randomNumber * (100 - 20 + 1)) + 20;
        const doc = new this.orderModel({
            ...body,
            price: randomInRange,
            distance: distance(),
        });
        const order = await doc.save();
        return order;
    }
    async getLowestPrice(body) {
        const findAll = await this.orderModel
            .find({ login: body.login })
            .sort({ price: 1 })
            .limit(1);
        if (findAll.length === 0) {
            throw new shared_1.UserDontHaveAnyOrders('U dont have any orders');
        }
        return findAll;
    }
    async getBiggestPrice(body) {
        const findAll = await this.orderModel.find({ login: body.login });
        if (findAll.length === 0) {
            throw new shared_1.UserDontHaveAnyOrders('U dont have any orders');
        }
        const highestOrder = findAll.reduce((maxOrder, currentOrder) => {
            return currentOrder.price > maxOrder.price ? currentOrder : maxOrder;
        }, findAll[0]);
        return highestOrder;
    }
    async getFiveLastFrom(body) {
        const last5From = [];
        const uniqueFromSet = new Set();
        const FiveLastFromPoints = await this.orderModel.find({ login: body.login }, { to: false });
        for (let i = FiveLastFromPoints.length - 1; i >= 0 && uniqueFromSet.size < 5; i--) {
            const currentFrom = FiveLastFromPoints[i].from;
            if (!uniqueFromSet.has(currentFrom)) {
                uniqueFromSet.add(currentFrom);
                console.log(uniqueFromSet);
                last5From.push(currentFrom);
            }
            console.log(last5From);
        }
        return last5From;
    }
    async getThreeLastTo(body) {
        const last3To = [];
        const uniqueFromSet = new Set();
        const FiveLastFromPoints = await this.orderModel.find({
            login: body.login,
        });
        for (let i = FiveLastFromPoints.length - 1; i >= 0 && uniqueFromSet.size < 5; i--) {
            const currentFrom = FiveLastFromPoints[i].to;
            if (!uniqueFromSet.has(currentFrom)) {
                uniqueFromSet.add(currentFrom);
                console.log(uniqueFromSet);
                last3To.push(currentFrom);
            }
        }
        return last3To;
    }
};
exports.OrderService = OrderService;
exports.OrderService = OrderService = __decorate([
    (0, common_1.Injectable)(),
    __param(0, (0, mongoose_2.InjectModel)(schema_1.Orders.name)),
    __param(1, (0, mongoose_2.InjectModel)(adresses_schema_1.Adresses.name)),
    __metadata("design:paramtypes", [mongoose_1.Model,
        mongoose_1.Model])
], OrderService);
//# sourceMappingURL=order.service.js.map