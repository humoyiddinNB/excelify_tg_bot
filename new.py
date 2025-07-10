import os
import asyncio
import logging
from typing import Dict, List
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ContextTypes, filters
import pandas as pd
from io import BytesIO
import tempfile

# Configure logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Bot token - replace with your actual bot token
BOT_TOKEN = "8015015567:AAEkxdcBMVtZe9J7BWLPY7zfAo68h3egLXI"

# Store user files temporarily
user_files: Dict[int, List[bytes]] = {}

class ExcelCombinerBot:
    def __init__(self):
        self.app = Application.builder().token(BOT_TOKEN).build()
        self.setup_handlers()
    
    def setup_handlers(self):
        """Set up command and message handlers"""
        self.app.add_handler(CommandHandler("start", self.start_command))
        self.app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx") | 
                                          filters.Document.FileExtension("xls"), 
                                          self.handle_excel_file))
        self.app.add_handler(CallbackQueryHandler(self.handle_callback))
    
    async def start_command(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Handle /start command"""
        user_id = update.effective_user.id
        user_files[user_id] = []  # Initialize empty file list for user
        
        welcome_message = """
🔗 **Excel File Combiner Bot**

Welcome! This bot helps you combine multiple Excel files into one.

**How to use:**
1. Send me your Excel files (.xlsx or .xls)
2. Click the "Combine Files" button when you're done uploading
3. I'll send you the combined Excel file

**Ready to start?** Send me your first Excel file!
        """
        
        await update.message.reply_text(
            welcome_message,
            parse_mode='Markdown'
        )
    
    async def handle_excel_file(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Handle Excel file uploads"""
        user_id = update.effective_user.id
        
        # Initialize user files if not exists
        if user_id not in user_files:
            user_files[user_id] = []
        
        try:
            # Get the file
            file = await update.message.document.get_file()
            file_data = await file.download_as_bytearray()
            
            # Store file data
            user_files[user_id].append(bytes(file_data))
            
            # Create keyboard with combine button
            keyboard = [
                [InlineKeyboardButton("📊 Combine Files", callback_data="combine")],
                [InlineKeyboardButton("🗑️ Clear All Files", callback_data="clear")]
            ]
            reply_markup = InlineKeyboardMarkup(keyboard)
            
            await update.message.reply_text(
                f"✅ **File received!** ({update.message.document.file_name})\n\n"
                f"📁 **Total files:** {len(user_files[user_id])}\n\n"
                f"Send more Excel files or click 'Combine Files' to merge them.",
                parse_mode='Markdown',
                reply_markup=reply_markup
            )
            
        except Exception as e:
            logger.error(f"Error handling file: {e}")
            await update.message.reply_text(
                "❌ Error processing your file. Please make sure it's a valid Excel file."
            )
    
    async def handle_callback(self, update: Update, context: ContextTypes.DEFAULT_TYPE):
        """Handle button callbacks"""
        query = update.callback_query
        await query.answer()
        
        user_id = query.from_user.id
        
        if query.data == "combine":
            await self.combine_files(query, user_id)
        elif query.data == "clear":
            await self.clear_files(query, user_id)
    
    async def combine_files(self, query, user_id: int):
        """Combine Excel files for the user"""
        if user_id not in user_files or not user_files[user_id]:
            await query.edit_message_text("❌ No files to combine! Please upload Excel files first.")
            return
        
        try:
            await query.edit_message_text("🔄 Combining your Excel files...")
            
            # Combine Excel files
            combined_data = []
            file_names = []
            
            for i, file_data in enumerate(user_files[user_id]):
                try:
                    # Read Excel file
                    df = pd.read_excel(BytesIO(file_data))
                    
                    # Add source file column
                    df['Source_File'] = f"File_{i+1}"
                    combined_data.append(df)
                    file_names.append(f"File_{i+1}")
                    
                except Exception as e:
                    logger.error(f"Error reading file {i+1}: {e}")
                    await query.edit_message_text(
                        f"❌ Error reading file {i+1}. Please make sure all files are valid Excel files."
                    )
                    return
            
            # Combine all dataframes
            if combined_data:
                combined_df = pd.concat(combined_data, ignore_index=True, sort=False)
                
                # Create output file
                output_buffer = BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    combined_df.to_excel(writer, sheet_name='Combined_Data', index=False)
                
                output_buffer.seek(0)
                
                # Send combined file
                await query.edit_message_text("📤 Sending combined file...")
                
                await query.message.reply_document(
                    document=output_buffer,
                    filename="combined_excel_file.xlsx",
                    caption=f"✅ **Combined Excel File**\n\n"
                           f"📊 **Total rows:** {len(combined_df)}\n"
                           f"📁 **Files combined:** {len(file_names)}\n"
                           f"📋 **Columns:** {len(combined_df.columns)}",
                    parse_mode='Markdown'
                )
                
                # Clear user files after successful combination
                user_files[user_id] = []
                
                await query.edit_message_text(
                    "✅ **Files combined successfully!**\n\n"
                    "Your combined Excel file has been sent above.\n"
                    "Send more files to create another combination!"
                )
            
        except Exception as e:
            logger.error(f"Error combining files: {e}")
            await query.edit_message_text(
                "❌ Error combining files. Please make sure all files are valid Excel files."
            )
    
    async def clear_files(self, query, user_id: int):
        """Clear all files for the user"""
        user_files[user_id] = []
        await query.edit_message_text(
            "🗑️ **All files cleared!**\n\n"
            "Send new Excel files to start over."
        )
    
    def run(self):
        """Start the bot"""
        print("🤖 Excel Combiner Bot is starting...")
        print("🔄 Bot is running! Press Ctrl+C to stop.")
        self.app.run_polling(allowed_updates=Update.ALL_TYPES)

def main():
    """Main function to run the bot"""
    if BOT_TOKEN == "YOUR_BOT_TOKEN_HERE":
        print("❌ Please set your bot token in BOT_TOKEN variable")
        print("💡 Get your bot token from @BotFather on Telegram")
        return
    
    bot = ExcelCombinerBot()
    bot.run()

if __name__ == "__main__":
    main()