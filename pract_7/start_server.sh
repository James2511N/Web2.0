for ((i = 1; i <= $1; i++)); do
    ((port=3000+i))
    PORT=$port node ./dist/server/server.js &
    sleep 5
done