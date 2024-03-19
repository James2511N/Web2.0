import { Router } from "express";
import { authorizationMiddleware } from "../middlewares.js";
import { ORDERS } from "../db.js";
import { ADDRESSES } from "../db.js";

export const OrdersRouter = Router();

const convertToDate = (date) => {
  /***
   * ^ -- початок рядка
   * \d -- перевірка на цифру
   * {N} -- N - разів повторень
   */
  // if (/^\d\d-(01|02|03|....|10|11|12)-\d{4}$/.test(query.createdAt)) { }
  if (!/^\d\d-\d\d-\d{4}$/.test(date)) {
    // return res.status(400).send({ message: `parameter createdAt has wrong format` });
    throw new Error(`parameter createdAt has wrong format`);
  }

  // const res = query.createdAt.split('-');
  // const month = res[1];
  const [day, month, year] = date.split("-");

  const mothsInt = parseInt(month);
  if (mothsInt < 1 || mothsInt > 12) {
    // return res.status(400).send({ message: `parameter createdAt has wrong month value` });

    throw new Error(`parameter createdAt has wrong month value`);
  }

  const result = new Date();
  result.setHours(2);
  result.setMinutes(0);
  result.setMilliseconds(0);
  result.setSeconds(0);

  result.setMonth(mothsInt - 1);
  result.setDate(day);
  result.setFullYear(year);

  return result;
};

const convertToDateMiddleware = (fieldName) => (req, res, next) => {
  const valueString = req.query[fieldName];

  if (!valueString) {
    return next();
  }
  try {
    const value = convertToDate(valueString);
    req.query[fieldName] = value;
    return next();
  } catch (err) {
    return res.status(400).send({ message: err.toString() });
  }
};

OrdersRouter.post("/orders", authorizationMiddleware, (req, res) => {
  const { from, to, type } = req.body;

  const isFromAddressValid = ADDRESSES.find((address) => address.name === from);
  const isToAddressValid = ADDRESSES.find((address) => address.name === to);

  if (type != "standard" && type != "lite" && type != "universal") {
    return res.status(400).send({ message: "Invalid type" });
  }

  if (!isFromAddressValid || !isToAddressValid) {
    return res.status(400).send({ message: "Invalid from or to address" });
  }
  const { body, user } = req;

  function distance() {
    const latitudeFrom = isFromAddressValid
      ? isFromAddressValid.location.latitude
      : null;
    const latitudeTo = isToAddressValid
      ? isToAddressValid.location.latitude
      : null;
    const longitudeFrom = isFromAddressValid
      ? isFromAddressValid.location.longitude
      : null;
    const longitudeTo = isToAddressValid
      ? isToAddressValid.location.longitude
      : null;
    var radlat1 = (Math.PI * latitudeFrom) / 180;
    var radlat2 = (Math.PI * latitudeTo) / 180;
    var theta = longitudeFrom - longitudeTo;
    var radtheta = (Math.PI * theta) / 180;
    var dist =
      Math.sin(radlat1) * Math.sin(radlat2) +
      Math.cos(radlat1) * Math.cos(radlat2) * Math.cos(radtheta);
    if (dist > 1) {
      dist = 1;
    }
    dist = Math.acos(dist);
    dist = (dist * 180) / Math.PI;
    dist = dist * 60 * 1.1515;
    return Math.round(dist * 1.609344);
  }

  const createdAt = new Date();
  createdAt.setHours(2);
  createdAt.setMinutes(0);
  createdAt.setMilliseconds(0);
  createdAt.setSeconds(0);

  function value() {
    if (type === "standard") {
      return distance() * 2.5;
    }
    if (type === "lite") {
      return distance() * 1.5;
    }
    if (type === "universal") {
      return distance() * 3;
    }
  }

  const order = {
    ...body,
    login: user.login,
    createdAt,
    status: "Active",
    id: crypto.randomUUID(),
    distance: distance(),
    value: value(),
  };

  ORDERS.push(order);

  return res.status(200).send({ message: "Order was created", order });
});

/**
 * GET /orders?createdAt=05-05-2024
 * GET /orders?createdAt= g mhdfbg kjdfbgkjd
 */
OrdersRouter.get(
  "/orders",
  authorizationMiddleware,
  convertToDateMiddleware("createdAt"),
  convertToDateMiddleware("createdFrom"),
  convertToDateMiddleware("createdTo"),
  (req, res) => {
    const { user, query } = req;

    if (query.createdAt && query.createdFrom && query.createdTo) {
      return res
        .status(400)
        .send({ message: "Too many parameter in query string" });
    }

    console.log(`query`, JSON.stringify(query));

    let orders = ORDERS.filter((el) => el.login === user.login);

    let active = ORDERS.filter((el) => el.status === "Active");

    function sortedOrders() {
      if (user.role == "Customer") {
        return orders;
      }
      if (user.role == "Admin") {
        return ORDERS;
      }
      if (user.role == "Driver") {
        return active;
      }
    }

    if (query.createdAt) {
      try {
        orders = ORDERS.filter((el) => {
          const value = new Date(el.createdAt);
          return value.getTime() === query.createdAt.getTime();
        });
      } catch (err) {
        return res.status(400).send({ message: err.toString() });
      }
    }

    if (query.createdFrom) {
      try {
        orders = ORDERS.filter((el) => {
          const value = new Date(el.createdAt);
          return value.getTime() >= query.createdFrom.getTime();
        });
      } catch (err) {
        return res.status(400).send({ message: err.toString() });
      }
    }

    if (query.createdTo) {
      try {
        orders = ORDERS.filter((el) => {
          const value = new Date(el.createdAt);
          return value.getTime() <= query.createdTo.getTime();
        });
      } catch (err) {
        return res.status(400).send({ message: err.toString() });
      }
    }

    return res.status(200).send(sortedOrders());
  }
);

/**
 * PATCH /orders/fhsdjkhfkdsj
 * PATCH /orders/fhsdjkhfkdsj12
 * PATCH /orders/fhsdjkhfkdsj123
 * PATCH /orders/fhsdjkhfkd123sj
 */

OrdersRouter.patch("/orders/:orderId", authorizationMiddleware, (req, res) => {
  const { params } = req;
  const { status } = req.body;
  const { user } = req;
  const { body } = req;

  let order = ORDERS.find((el) => el.id === params.orderId);

  if (!order) {
    return res
      .status(400)
      .send({ message: `Order with id ${params.orderId} was not found` });
  }
  if (order.login != user.login && user.role != "Admin") {
    return res.status(400).send({ message: "Invalid login" });
  }
  if (
    body.status != "Rejected" &&
    body.status != "In progress" &&
    body.status != "Done"
  ) {
    return res.status(400).send({ message: "Invalid statusss" });
  }

    if (order.login == user.login && user.role == "Customer") {
      if (body.status == "Rejected") {
        ORDERS.update((el) => el.id === params.orderId, {
          status: body.status,
        });
      }
      if (body.status != "Rejected") {
        return res.status(400).send({ message: "Invalid status" });
      }
    }
  if (user.role == "Admin" && body.status == "Rejected") {
    if (order.status == "Active") {
      ORDERS.update((el) => el.id === params.orderId, { status: body.status });
    }
    if (order.status != "Active") {
      return res.status(400).send({ message: "Invalid" });
    }
  }
  if (user.role == "Admin" && body.status == "In progress") {
    if (order.status == "Rejected") {
      ORDERS.update((el) => el.id === params.orderId, { status: body.status });
    }
    if (order.status != "Rejected") {
      return res.status(400).send({ message: "Invalid" });
    }
  }
  if (user.role == "Admin" && body.status == "Done") {
    if (order.status == "In progress") {
      ORDERS.update((el) => el.id === params.orderId, { status: body.status });
    }
    if (order.status != "In progress") {
      return res.status(400).send({ message: "Invalid" });
    }
  }
  if (user.role == "Driver" && body.status == "Active") {
    if (order.status == "In progress") {
      ORDERS.update((el) => el.id === params.orderId, { status: body.status });
    }
    if (order.status != "In progress") {
      return res.status(400).send({ message: "Invalid" });
    }
  }
  if (user.role == "Driver" && body.status == "In progress") {
    if (order.status == "Done") {
      ORDERS.update((el) => el.id === params.orderId, { status: body.status });
    }
    if (order.status != "Done") {
      return res.status(400).send({ message: "Invalid" });
    }
  }
  if (
    user.role == "Driver" &&
    body.status != "In progress" &&
    body.status != "Active"
  ) {
    return res.status(400).send({ message: "Invalid access" });
  }

  order = ORDERS.find((el) => el.id === params.orderId);
  return res.status(200).send(order);
});
