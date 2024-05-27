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
Object.defineProperty(exports, "__esModule", { value: true });
exports.AdressesSchema = exports.Adresses = void 0;
const mongoose_1 = require("@nestjs/mongoose");
let Adresses = class Adresses {
};
exports.Adresses = Adresses;
__decorate([
    (0, mongoose_1.Prop)({ type: String, required: true }),
    __metadata("design:type", String)
], Adresses.prototype, "name", void 0);
__decorate([
    (0, mongoose_1.Prop)({
        type: Object,
        required: true,
        location: {
            longitude: { type: Number, required: true },
            latitude: { type: Number, required: true },
        },
    }),
    __metadata("design:type", Object)
], Adresses.prototype, "location", void 0);
exports.Adresses = Adresses = __decorate([
    (0, mongoose_1.Schema)({ collection: 'adresses' })
], Adresses);
exports.AdressesSchema = mongoose_1.SchemaFactory.createForClass(Adresses);
//# sourceMappingURL=adresses.schema.js.map