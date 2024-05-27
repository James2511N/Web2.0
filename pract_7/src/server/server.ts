import { NestFactory } from '@nestjs/core';
import { ServerModule } from './server.module';
import { ConfigService } from '@nestjs/config';
import { getRandomTime } from './shared/getRandomTime';

async function bootstrap() {
  const app = await NestFactory.create(ServerModule);
  const configService = app.get(ConfigService);
  const port = configService.get('PORT');

  await app.listen(port);


  const minTime = 500
  const maxTime = 1000
  const time = getRandomTime(maxTime, minTime)
  console.log(`Server will start on ${port} and shutdown after ${time}s`)
  setTimeout(() => {
    app.close()
  }, time * 1000)
}
bootstrap();
