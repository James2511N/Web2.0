import { Injectable } from '@nestjs/common';
import { Telegraf } from 'telegraf';
import { Context } from 'telegraf/typings/context';
import { Update } from 'telegraf/typings/core/types/typegram';
import { InjectModel } from '@nestjs/mongoose';
import { Model } from 'mongoose';
import { Link } from '../link/link.model';

@Injectable()
export class BotService {
  private bot: Telegraf<Context<Update>>;

  constructor(@InjectModel('Link') private readonly linkModel: Model<Link>) {
    this.bot = new Telegraf('-');

    this.bot.start((ctx) => ctx.reply('Welcome!'));

    this.bot.command('addlink', async (ctx) => {
      const userId = ctx.from.id;
      const url = ctx.message.text.split(' ')[1];
      if (!url) {
        return ctx.reply('Будь ласка, надайте URL.');
      }
      const existingLink = await this.linkModel.findOne({ url, userId });
      if (existingLink) {
        return ctx.reply('Ви вже слідкуєте за цим посиланням.');
      }
      const newLink = new this.linkModel({ url, userId });
      await newLink.save();
      ctx.reply('Посилання успішно додано!');
    });

    this.bot.command('links', async (ctx) => {
      const userId = ctx.from.id;
      const links = await this.linkModel.find({ userId });

      links.forEach((link) => {
        ctx.reply(link.url, {
          reply_markup: {
            inline_keyboard: [
              [
                {
                  text: '❌',
                  callback_data: `delete_${link._id}`,
                },
              ],
            ],
          },
        });
      });
    });

    this.bot.action(/delete_(.+)/, async (ctx) => {
      const linkId = ctx.match[1];
      const link = await this.linkModel.findById(linkId);

      if (link.userId === ctx.from.id) {
        await this.linkModel.deleteOne({ _id: linkId });
        ctx.answerCbQuery('Посилання видалено!');
      } else {
        ctx.answerCbQuery('Нема прав');
      }
    });

    this.bot.launch();
  }
}
