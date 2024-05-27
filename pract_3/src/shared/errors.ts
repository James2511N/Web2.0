export class UserAlreadyExists extends Error {
  constructor(msg: string) {
    super(msg);
  }
}

export class UserNotFound extends Error {
  constructor(msg: string) {
    super(msg);
  }
}

export class UserDontHaveAnyOrders extends Error {
  constructor(msg: string) {
    super(msg);
  }
}
