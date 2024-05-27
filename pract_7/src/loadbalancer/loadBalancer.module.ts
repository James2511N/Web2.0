import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { HttpModule } from '@nestjs/axios';
import { LoadBalancerController } from './loadBalancer.controller';
import { LoadBalancerService } from './loadBalancer.service';
import { Server, ServerSchema } from 'src/schemas/server.schema';

@Module({
  imports: [
    MongooseModule.forRoot('mongodb+srv://mynameseriy:gpyJZTvG6rY7r0wg@cluster0.brv3ly5.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0'),
    MongooseModule.forFeature([{ name: Server.name, schema: ServerSchema }]),
    HttpModule,
  ],
  controllers: [LoadBalancerController],
  providers: [LoadBalancerService],
})
export class LoadBalancerModule {}

