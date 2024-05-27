# Instruction

1. After each change before start server you will need to run `npm run build`
2. On server you need to implement next APIs
    * GET /api/metrics -- should return information about RAM and CPU
    * GET /api/health -- should always respond with status 200
    * POST /orders -- already implemented
3. On balancer server you need to implement:
    * connection to db. db name should be balancer
        ** create collection with structure that we have in Scale.drawio.png
    * POST /instances -- will be used when new server will start. This API should register new instance in collections servers with status active
    * GET /instances -- should return all active instances
    * We should get metrics from each active server one time per minute and save them in collection
    * we should execute health check each 30s for each active server
        ** We should change status of server in the case if we get 3 fails of heal check
    * when we get any other request we should redirect it to one of the servers with state active
        ** Select server where we have minimum CPU and RAM usage

# start server

```terminal:git-bash
bash start_server.sh COUNT
```
COUNT -- number of servers