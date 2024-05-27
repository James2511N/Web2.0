/// <reference types="mongoose/types/aggregate" />
/// <reference types="mongoose/types/callback" />
/// <reference types="mongoose/types/collection" />
/// <reference types="mongoose/types/connection" />
/// <reference types="mongoose/types/cursor" />
/// <reference types="mongoose/types/document" />
/// <reference types="mongoose/types/error" />
/// <reference types="mongoose/types/expressions" />
/// <reference types="mongoose/types/helpers" />
/// <reference types="mongoose/types/middlewares" />
/// <reference types="mongoose/types/indexes" />
/// <reference types="mongoose/types/models" />
/// <reference types="mongoose/types/mongooseoptions" />
/// <reference types="mongoose/types/pipelinestage" />
/// <reference types="mongoose/types/populate" />
/// <reference types="mongoose/types/query" />
/// <reference types="mongoose/types/schemaoptions" />
/// <reference types="mongoose/types/schematypes" />
/// <reference types="mongoose/types/session" />
/// <reference types="mongoose/types/types" />
/// <reference types="mongoose/types/utility" />
/// <reference types="mongoose/types/validation" />
/// <reference types="mongoose/types/virtuals" />
/// <reference types="mongoose/types/inferschematype" />
import { Types, Document } from 'mongoose';
export declare class Orders {
    from: string;
    to: string;
    price: number;
    login: string;
    type: string;
    distance: string;
}
export declare const OrdersSchema: import("mongoose").Schema<Orders, import("mongoose").Model<Orders, any, any, any, Document<unknown, any, Orders> & Orders & {
    _id: Types.ObjectId;
}, any>, {}, {}, {}, {}, import("mongoose").DefaultSchemaOptions, Orders, Document<unknown, {}, import("mongoose").FlatRecord<Orders>> & import("mongoose").FlatRecord<Orders> & {
    _id: Types.ObjectId;
}>;
export type OrdersLeanDoc = Orders & {
    _id: Types.ObjectId;
};
export type OrdersDoc = Orders & Document;
