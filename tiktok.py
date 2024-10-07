from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackContext
from telegram.ext import filters
from openpyxl import Workbook
import os
import asyncio

YOUR_BOT_TOKEN = os.getenv("YOUR_BOT_TOKEN")  # Sử dụng token của bạn

# Tạo file Excel nếu không tồn tại
excel_file = "messages.xlsx"
if not os.path.exists(excel_file):
    wb = Workbook()
    wb.save(excel_file)

user_status = {}
user_messages = {}
user_exported_index = {}
user_timers = {}

async def export_user_messages(update, context, user_id):
    user_file = f"{user_id}_messages.xlsx"
    new_wb = Workbook()
    new_ws = new_wb.active
    new_ws.title = str(user_id)
    new_ws.append(["User ID", "Message"])

    if user_id in user_messages:
        for msg in user_messages[user_id][user_exported_index[user_id]:]:
            new_ws.append([user_id, msg])

    new_wb.save(user_file)

    with open(user_file, 'rb') as file:
        await context.bot.send_document(chat_id=update.message.chat.id, document=file)

async def stop_bot(update: Update, context: CallbackContext, user_id):
    await asyncio.sleep(1800)  # Thay đổi thời gian chờ thành 60 giây
    if user_id in user_status and user_status[user_id]:
        if user_exported_index[user_id] < len(user_messages[user_id]):
            await export_user_messages(update, context, user_id)
        await update.message.reply_text("Bot sẽ dừng do không hoạt động trong 30 phút.")
        user_status[user_id] = False
        user_messages[user_id] = []

async def start(update: Update, context: CallbackContext) -> None:
    user_id = update.message.from_user.id
    user_status[user_id] = True
    user_messages[user_id] = []
    user_exported_index[user_id] = 0
    await update.message.reply_text('Xin chào! Tôi sẽ lưu tin nhắn của bạn vào file Excel từ bây giờ. Bắt đầu gửi tin nhắn!')

    if user_id in user_timers:
        user_timers[user_id].cancel()
    user_timers[user_id] = asyncio.create_task(stop_bot(update, context, user_id))

async def thongtin(update: Update, context: CallbackContext) -> None:
    commands_info = (
        "/start - Bắt đầu lưu tin nhắn.",
        "/export - Xuất tin nhắn đã lưu.",
        "/thongtin - Hiển thị thông tin về các lệnh hỗ trợ."
    )
    await update.message.reply_text("\n".join(commands_info))

async def export(update: Update, context: CallbackContext) -> None:
    user_id = update.message.from_user.id
    await export_user_messages(update, context, user_id)
    user_exported_index[user_id] = len(user_messages[user_id])

async def echo(update: Update, context: CallbackContext) -> None:
    user_id = update.message.from_user.id
    message_text = update.message.text

    if user_status.get(user_id, False):
        user_messages[user_id].append(message_text)
        await update.message.reply_text(f'Tin nhắn của bạn đã được lưu: "{message_text}"')

        if user_id in user_timers:
            user_timers[user_id].cancel()
        user_timers[user_id] = asyncio.create_task(stop_bot(update, context, user_id))
    else:
        await update.message.reply_text("Vui lòng sử dụng lệnh /start để bắt đầu lưu tin nhắn.")

def main() -> None:
    application = ApplicationBuilder().token(YOUR_BOT_TOKEN).build()
    
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("thongtin", thongtin))
    application.add_handler(CommandHandler("export", export))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, echo))

    application.run_polling()

if __name__ == '__main__':
    main()