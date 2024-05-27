export const getRandomTime = (maxTime: number, minTime: number) =>{
    return Math.floor(Math.random() * maxTime) + minTime
}