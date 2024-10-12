from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackContext
from telegram.ext import filters
from openpyxl import Workbook, load_workbook
import os
import asyncio
import requests
from bs4 import BeautifulSoup

YOUR_BOT_TOKEN = os.getenv("YOUR_BOT_TOKEN")  # Sử dụng token của bạn

# Cấu hình proxy
proxy = {
    'http': 'http://d3530cadb9:M91VxFDm@130.44.202.141:4444',
    'https': 'http://d3530cadb9:M91VxFDm@130.44.202.141:4444',
}

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
    await asyncio.sleep(1800)  # Thời gian chờ 30 phút
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
        "/thongtin - Hiển thị thông tin về các lệnh hỗ trợ.",
        "/readfile - Đọc file Excel và lấy dữ liệu sản phẩm."
    )
    await update.message.reply_text("\n".join(commands_info))

async def export(update: Update, context: CallbackContext) -> None:
    user_id = update.message.from_user.id
    await export_user_messages(update, context, user_id)
    user_exported_index[user_id] = len(user_messages[user_id])

async def fetch_url_data(url):
    try:
        response = requests.get(url, proxies=proxy)  # Sử dụng proxy
        response.raise_for_status()  # Kiểm tra mã trạng thái

        soup = BeautifulSoup(response.content, 'html.parser')

        # Lấy tiêu đề từ thẻ <div> có class 'index-title--AnTxK'
        title_div = soup.find('div', class_='index-title--AnTxK')
        title = title_div.get_text() if title_div else "Không tìm thấy tiêu đề."

        # Kiểm tra sự tồn tại của slick-track và lấy hình ảnh
        slick_track = soup.find('div', class_='slick-track')
        img_tags = slick_track.find_all('img') if slick_track else []
        img_urls = [img.get('src') for img in img_tags if img.get('src')]

        return title, img_urls

    except requests.exceptions.RequestException as e:
        return "Không lấy được dữ liệu", []

async def echo(update: Update, context: CallbackContext) -> None:
    user_id = update.message.from_user.id
    message_text = update.message.text

    if user_status.get(user_id, False):
        user_messages[user_id].append(message_text)
        await update.message.reply_text(f'Tin nhắn của bạn đã được lưu: "{message_text}"')

        # Kiểm tra nếu tin nhắn là một URL
        if message_text.startswith('http://') or message_text.startswith('https://'):
            title, img_urls = await fetch_url_data(message_text)
            await update.message.reply_text(f'Tiêu đề: {title}\nSố lượng hình ảnh: {len(img_urls)}')
        else:
            await update.message.reply_text("Đó không phải là một URL hợp lệ.")

        # Thiết lập lại thời gian chờ
        if user_id in user_timers:
            user_timers[user_id].cancel()
        user_timers[user_id] = asyncio.create_task(stop_bot(update, context, user_id))
    else:
        await update.message.reply_text("Vui lòng sử dụng lệnh /start để bắt đầu lưu tin nhắn.")

async def read_excel_file(file_path):
    """Đọc file Excel và lấy các link sản phẩm từ cột 'Link'."""
    df = load_workbook(file_path)
    results = []

    for sheet in df.sheetnames:
        worksheet = df[sheet]
        for row in worksheet.iter_rows(min_row=2, values_only=True):  # Bỏ qua hàng tiêu đề
            url = row[0]  # Giả sử link nằm ở cột đầu tiên
            title, img_urls = await fetch_url_data(url)
            results.append({'URL': url, 'Title': title, 'Images': img_urls})

    return results

async def read_file(update: Update, context: CallbackContext) -> None:
    user_id = update.message.from_user.id

    # Kiểm tra trạng thái người dùng đã bắt đầu tương tác
    if not user_status.get(user_id, False):
        await update.message.reply_text("Vui lòng sử dụng lệnh /start để bắt đầu.")
        return

    # Kiểm tra nếu có file được gửi
    if update.message.document:
        file = await update.message.document.get_file()
        input_file = f"{user_id}_input_file.xlsx"
        await file.download_to_drive(input_file)  # Tải file về

        results = await read_excel_file(input_file)

        # Tạo file Excel mới với kết quả
        output_file = f"{user_id}_output.xlsx"
        output_wb = Workbook()
        output_ws = output_wb.active
        output_ws.append(["URL", "Title"] + [f"Image URL {i+1}" for i in range(max(len(result['Images']) for result in results))])  # Tiêu đề cột

        for result in results:
            row = [result['URL'], result['Title']]
            row.extend(result['Images'])  # Thêm từng link hình ảnh vào hàng
            output_ws.append(row)

        output_wb.save(output_file)

        with open(output_file, 'rb') as f:
            await context.bot.send_document(chat_id=update.message.chat.id, document=f)

        await update.message.reply_text("File đã được xử lý. Bạn có thể gửi file mới bất kỳ lúc nào.")
    else:
        await update.message.reply_text("Vui lòng gửi file Excel để xử lý.")

def main() -> None:
    application = ApplicationBuilder().token(YOUR_BOT_TOKEN).build()
    
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("thongtin", thongtin))
    application.add_handler(CommandHandler("export", export))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, echo))
    application.add_handler(MessageHandler(filters.Document.ALL, read_file))  # Xử lý file tải lên

    application.run_polling()

if __name__ == '__main__':
    main()