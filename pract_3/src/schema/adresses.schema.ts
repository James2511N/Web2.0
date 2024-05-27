import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose'
import { Types, Document } from 'mongoose'

@Schema({ collection: 'adresses' })
export class Adresses {
    @Prop({ type: String, required: true })
    name: string

    @Prop({
        type: Object,
        required: true,
        location: {
            longitude: { type: Number, required: true },
            latitude: { type: Number, required: true },
        },
    })
    location: {
        longitude: number
        latitude: number
    }
}

export const AdressesSchema = SchemaFactory.createForClass(Adresses)

export type AdressesLeanDoc = Adresses & { _id: Types.ObjectId }
export type AdressesDoc = Adresses & Document
