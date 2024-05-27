import { Injectable } from '@nestjs/common'
import { OrderDto } from '../models'
import { Orders, OrdersDoc } from '../schema'
import { Adresses, AdressesDoc } from '../schema/adresses.schema'
import { Model } from 'mongoose'
import { InjectModel } from '@nestjs/mongoose'
import { UserDto } from '../models'
import { UserDontHaveAnyOrders } from 'src/shared'

@Injectable()
export class OrderService {
    constructor(
        @InjectModel(Orders.name)
        private readonly orderModel: Model<OrdersDoc>,
        @InjectModel(Adresses.name)
        private readonly addressModel: Model<AdressesDoc>
    ) { }

    async createOrder(body: OrderDto & { login: string }) {
        const fromAddressExists = await this.addressModel.exists({
            name: body.from,
        })
        const toAddressExists = await this.addressModel.exists({
            name: body.to,
        })
        if (!fromAddressExists || !toAddressExists) {
            throw new Error(
                'Either the From or To address is invalid and not found in our records.'
            )
        }

        const fromAddressDocument = await this.addressModel.findOne({ name: body.from });
        const toAddressDocument = await this.addressModel.findOne({ name: body.to });


        function distance() {
            const latitudeFrom = fromAddressDocument
                ? fromAddressDocument.location.latitude
                : null
            const latitudeTo = toAddressDocument
                ? toAddressDocument.location.latitude
                : null
            const longitudeFrom = fromAddressDocument
                ? fromAddressDocument.location.longitude
                : null
            const longitudeTo = toAddressDocument
                ? toAddressDocument.location.longitude
                : null
            var radlat1 = (Math.PI * latitudeFrom) / 180
            var radlat2 = (Math.PI * latitudeTo) / 180
            var theta = longitudeFrom - longitudeTo
            var radtheta = (Math.PI * theta) / 180
            var dist =
                Math.sin(radlat1) * Math.sin(radlat2) +
                Math.cos(radlat1) * Math.cos(radlat2) * Math.cos(radtheta)
            if (dist > 1) {
                dist = 1
            }
            dist = Math.acos(dist)
            dist = (dist * 180) / Math.PI
            dist = dist * 60 * 1.1515
            return Math.round(dist * 1.609344)
        }
        // function value() {
        //   if (type === "standard") {
        //     return distance() * 2.5;
        //   }
        //   if (type === "lite") {
        //     return distance() * 1.5;
        //   }
        //   if (type === "universal") {
        //     return distance() * 3;
        //   }
        // }
        const randomNumber = Math.random()
        const randomInRange = Math.floor(randomNumber * (100 - 20 + 1)) + 20

        const doc = new this.orderModel({
            ...body,
            price: randomInRange,
            distance: distance(),
        })

        const order = await doc.save()

        return order
    }

    async getLowestPrice(body: UserDto) {
        const findAll = await this.orderModel
            .find({ login: body.login })
            .sort({ price: 1 })
            .limit(1) // price:1 сортує від найнижчого до найбільшого і видає 1 результат
        if (findAll.length === 0) {
            throw new UserDontHaveAnyOrders('U dont have any orders')
        }
        return findAll
    }

    async getBiggestPrice(body: UserDto) {
        const findAll = await this.orderModel.find({ login: body.login })
        if (findAll.length === 0) {
            throw new UserDontHaveAnyOrders('U dont have any orders')
        }
        const highestOrder = findAll.reduce((maxOrder, currentOrder) => {
            return currentOrder.price > maxOrder.price ? currentOrder : maxOrder
        }, findAll[0])

        return highestOrder
    }
    async getFiveLastFrom(body: UserDto) {
        const last5From = []
        const uniqueFromSet = new Set()
        const FiveLastFromPoints = await this.orderModel.find(
            { login: body.login },

            { to: false }
        )
        for (
            let i = FiveLastFromPoints.length - 1;
            i >= 0 && uniqueFromSet.size < 5;
            i--
        ) {
            const currentFrom = FiveLastFromPoints[i].from
            // console.log(currentTo)
            if (!uniqueFromSet.has(currentFrom)) {
                uniqueFromSet.add(currentFrom)
                console.log(uniqueFromSet)
                last5From.push(currentFrom)
            }

            console.log(last5From)
        }
        return last5From
    }
    async getThreeLastTo(body: UserDto) {
        const last3To = []
        const uniqueFromSet = new Set()
        const FiveLastFromPoints = await this.orderModel.find({
            login: body.login,
        })
        for (
            let i = FiveLastFromPoints.length - 1;
            i >= 0 && uniqueFromSet.size < 5;
            i--
        ) {
            const currentFrom = FiveLastFromPoints[i].to
            // console.log(currentTo)
            if (!uniqueFromSet.has(currentFrom)) {
                uniqueFromSet.add(currentFrom)
                console.log(uniqueFromSet)
                last3To.push(currentFrom)
            }
        }
        return last3To
    }
}
