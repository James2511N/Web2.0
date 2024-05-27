"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.UserDontHaveAnyOrders = exports.UserNotFound = exports.UserAlreadyExists = void 0;
class UserAlreadyExists extends Error {
    constructor(msg) {
        super(msg);
    }
}
exports.UserAlreadyExists = UserAlreadyExists;
class UserNotFound extends Error {
    constructor(msg) {
        super(msg);
    }
}
exports.UserNotFound = UserNotFound;
class UserDontHaveAnyOrders extends Error {
    constructor(msg) {
        super(msg);
    }
}
exports.UserDontHaveAnyOrders = UserDontHaveAnyOrders;
//# sourceMappingURL=errors.js.map