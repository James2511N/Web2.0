import express from "express"
import bodyParser from "body-parser"
import { USERS, ORDERS } from "./db.js"
import { authorizationMiddleware } from "./middlewares.js"

const app = express()

app.use(bodyParser.json())

/**
 * POST -- create resource
 * req -> input data
 * res -> output data
 */
app.post("/users", (req, res) => {
  const { body } = req

  console.log(`body`, JSON.stringify(body))

  const isUserExist = USERS.some((el) => el.login === body.login)
  if (isUserExist) {
    return res
      .status(400)
      .send({ message: `user with login ${body.login} already exists` })
  }

  USERS.push(body)

  res.status(200).send({ message: "User was created" })
})

app.get("/users", (req, res) => {
  const users = USERS.map((user) => {
    const { password, ...other } = user
    return other
  })
  return res.status(200).send(users)
})

app.post("/login", (req, res) => {
  const { body } = req

  const user = USERS.find(
    (el) => el.login === body.login && el.password === body.password
  )

  if (!user) {
    return res.status(400).send({ message: "User was not found" })
  }

  const token = crypto.randomUUID()

  user.token = token
  USERS.save(user.login, { token })

  return res.status(200).send({
    token,
    message: "User was login",
  })
})

app.post("/orders", authorizationMiddleware, (req, res) => {
  const { body, user } = req
  const randomNumber = Math.random()
  const randomInRange = Math.floor(randomNumber * (100 - 20 + 1)) + 20

  const order = {
    ...body,
    login: user.login,
    price: randomInRange,
  }

  ORDERS.push(order)

  return res.status(200).send({ message: "Order was created", order })
})

app.get("/orders", authorizationMiddleware, (req, res) => {
  const { user } = req

  const orders = ORDERS.filter((el) => el.login === user.login)

  return res.status(200).send(orders)
})

app.get("/address/from/last-5", authorizationMiddleware, (req, res) => {
  try {
    const { user } = req
    const orders = ORDERS.filter((order) => order.login === user.login)
    const uniqueFromAddresses = Array.from(
      new Set(orders.map((order) => order.from))
    )
    const last5UniqueFromAddresses = uniqueFromAddresses.slice(-5)
    console.log(last5UniqueFromAddresses)
    return res.status(200).json(last5UniqueFromAddresses)
  } catch (error) {
    console.error(error)
    return res.status(500).json({ message: "Internal Server Error" })
  }
})

// array - переводить в ерей
// set - масив в якому унікальні обєкти ключ-значення

app.get("/address/to/last-3", authorizationMiddleware, (req, res) => {
  try {
    const { user } = req
    const orders = ORDERS.filter((order) => order.login === user.login)
    const uniqueFromAddresses = Array.from(
      new Set(orders.map((order) => order.to))
    )

    const last5UniqueToAddresses = uniqueFromAddresses.slice(-5)
    console.log(last5UniqueToAddresses)
    return res.status(200).json(last5UniqueToAddresses)
  } catch (error) {
    console.error(error)
    return res.status(500).json({ message: "Internal Server Error" })
  }
})

app.get("/orders/lowest", authorizationMiddleware, (req, res) => {
  try {
    const { user } = req
    const orders = ORDERS.filter((order) => order.login === user.login)

    if (orders.length === 0) {
      return res.status(404).send({ message: "User do not have orders yet" })
    }
    const lowestOrder = orders.reduce((minOrder, currentOrder) => {
      return currentOrder.price < minOrder.price ? currentOrder : minOrder
    })

    return res.status(200).json(lowestOrder)
  } catch (error) {
    console.error(error)
    return res.status(500).json({ message: "Internal Server Error" })
  }
})

app.get("/orders/biggest", authorizationMiddleware, (req, res) => {
  try {
    const { user } = req
    const orders = ORDERS.filter((order) => order.login === user.login)

    if (orders.length === 0) {
      return res.status(404).send({ message: "User do not have orders yet" })
    }
    const biggestOrder = orders.reduce((minOrder, currentOrder) => {
      return currentOrder.price > minOrder.price ? currentOrder : minOrder
    })

    return res.status(200).json(biggestOrder)
  } catch (error) {
    console.error(error)
    return res.status(500).json({ message: "Internal Server Error" })
  }
})

app.listen(8080, () => console.log("Server was started"))
