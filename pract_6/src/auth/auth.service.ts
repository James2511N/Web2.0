import { Injectable, UnauthorizedException } from '@nestjs/common';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { User } from './schemas/user.schema';
import * as bcrypt from 'bcryptjs';
import { CreateUserDto } from './dto/create-user.dto';
import { LoginUserDto } from './dto/login-user.dto';

@Injectable()
export class AuthService {
  constructor(
    @InjectModel('User') private readonly userModel: Model<User>
  ) {}

  async createUser(createUserDto: CreateUserDto): Promise<User> {
    const { username, password } = createUserDto;
    const hashedPassword = await bcrypt.hash(password, 10);
    const user = new this.userModel({ username, password: hashedPassword });
    return user.save();
  }

  async validateUser(username: string, password: string): Promise<any> {
    const user = await this.userModel.findOne({ username });
    if (user && bcrypt.compareSync(password, user.password)) {
      const { password, ...result } = user;
      return result;
    }
    return null;
  }
}
