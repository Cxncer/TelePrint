import os
import logging
import tempfile
import subprocess
import time
import asyncio
import json
from datetime import datetime, date
from typing import Dict, Optional
from dataclasses import dataclass, asdict
from enum import Enum
import html
from dotenv import load_dotenv

from pypdf import PdfReader
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)
from telegram.constants import ParseMode

# Load environment variables
load_dotenv()

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
BOT_TOKEN = os.getenv("BOT_TOKEN")
STAFF_GROUP_ID = int(os.getenv("STAFF_GROUP_ID", "0"))
TARGET_THREAD_ID = int(os.getenv("TARGET_THREAD_ID", "0"))
PRINTER_NAME = os.getenv("PRINTER_NAME")
PRINTER_IP = os.getenv("PRINTER_IP")

SUMATRA_PATH = r"C:\Users\Noble 2\AppData\Local\SumatraPDF\SumatraPDF.exe"

DOWNLOAD_DIR = os.path.join(tempfile.gettempdir(), "printshop_jobs")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

ALLOWED_EXT = [".pdf", ".jpg", ".jpeg", ".png"]

# Job limits
MAX_COPIES = 100
MAX_PAGES = 500
MIN_COPIES = 1

# Print timeout in seconds
PRINT_TIMEOUT = 60

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# KHMER LANGUAGE TRANSLATIONS
# ─────────────────────────────────────────────
KHMER = {
    # Buttons
    "a4": "📄 A4",
    "a3": "📃 A3",
    "color": "🎨 ពណ៌",
    "bw": "⚫ ខ្មៅស",
    "single": "📄 ម្ខាង",
    "double": "📄📄 សងខាង",
    "cancel": "❌ បោះបង់",
    "custom": "✏️ បញ្ចូលដោយខ្លួនឯង",
    
    # Messages
    "select_paper": "*ជ្រើសរើសទំហំក្រដាស:*",
    "select_color": "*ជ្រើសរើសពណ៌:*",
    "select_duplex": "*ជ្រើសរើសការបោះពុម្ព:*",
    "select_copies": "*ជ្រើសរើសចំនួនច្បាប់:*",
    "enter_copies": "*បញ្ចូលចំនួនច្បាប់:*",
    "job": "🖨️ *ការងារលេខ #{job_number}*",
    "file": "📄 `{file_name}`",
    "user": "👤 អ្នកប្រើប្រាស់: {username}",
    "completed": "✅ *ការងារលេខ #{job_number} បានបញ្ចប់!*",
    "pages": "📄 ទំព័រ: {pages}",
    "copies": "📄 ច្បាប់: {copies}",
    "size": "📄 ទំហំ: {paper_size}",
    "price": "💰 តម្លៃ: *{price:,} រៀល*",
    "printed": "🕐 បោះពុម្ពនៅ: {time}",
    
    # Errors
    "unsupported_file": "❌ ឯកសារមិនគាំទ្រ។ សូមផ្ញើ PDF, JPG, ឬ PNG",
    "download_failed": "❌ ទាញយកឯកសារមិនបានសម្រេច",
    "print_failed": "❌ ការបោះពុម្ពបរាជ័យ",
    "wrong_group": "❌ បូតនេះប្រើបានតែក្នុងក្រុមដែលបានកំណត់",
    "wrong_thread": "❌ សូមប្រើខ្សែសន្ទនាដែលបានកំណត់",
    "expired": "⚠️ ការងារនេះបានផុតកំណត់ ឬត្រូវបានបោះបង់",
    "cancelled": "❌ ការងារលេខ #{job_number} ត្រូវបានបោះបង់",
}

def translate(key: str, **kwargs) -> str:
    """Get Khmer translation with formatting"""
    text = KHMER.get(key, key)
    if kwargs:
        return text.format(**kwargs)
    return text

# Helper function to escape Markdown
def escape_markdown(text):
    """Escape Markdown special characters"""
    if not text:
        return ""
    special_chars = ['_', '*', '[', ']', '(', ')', '~', '`', '>', '#', '+', '-', '=', '|', '{', '}', '.', '!']
    for char in special_chars:
        text = text.replace(char, f'\\{char}')
    return text

# ─────────────────────────────────────────────
# THREAD ACCESS CHECK (SILENT)
# ─────────────────────────────────────────────
async def check_thread_access(update: Update, silent: bool = True) -> bool:
    """
    Check if message is in the correct thread
    If silent=True, don't send error messages (just return False)
    """
    msg = update.message or (update.callback_query and update.callback_query.message)
    
    if not msg:
        return False
    
    # Check group
    if msg.chat.id != STAFF_GROUP_ID:
        if not silent and update.message:
            await update.message.reply_text(translate("wrong_group"))
        return False
    
    # Check thread/topic - silently ignore wrong threads
    if TARGET_THREAD_ID and msg.message_thread_id != TARGET_THREAD_ID:
        return False  # Silent ignore - no response
    
    return True

# ─────────────────────────────────────────────
# DEBUGGING COMMANDS
# ─────────────────────────────────────────────
async def ping(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Simple ping command to test if bot is responding"""
    if not await check_thread_access(update, silent=False):
        return
    await update.message.reply_text("🏓 Pong! Bot is alive and responding!")

async def test_duplex(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Test different duplex settings to find what works"""
    if not await check_thread_access(update, silent=False):
        return
        
    if not update.message or not update.message.reply_to_message:
        await update.message.reply_text(
            "❌ Please reply to a document with /test_duplex\n\n"
            "How to use:\n"
            "1. Send a document to the bot\n"
            "2. Tap 'Reply' on that document message\n"
            "3. Type /test_duplex and send"
        )
        return
    
    msg = update.message.reply_to_message
    
    # Check if we already have this file in a job
    file_path = None
    for job_id, job in job_manager.jobs.items():
        if str(msg.message_id) in job_id:
            file_path = job.file_path
            break
    
    # If not found, try to get from replied message
    if not file_path and msg.document:
        file = msg.document
        file_path = os.path.join(DOWNLOAD_DIR, f"duplex_test_{file.file_name}")
        
        status_msg = await update.message.reply_text("📥 Downloading file...")
        try:
            tg_file = await context.bot.get_file(file.file_id)
            await tg_file.download_to_drive(file_path)
            await status_msg.delete()
        except Exception as e:
            await status_msg.edit_text(f"❌ Download failed: {str(e)}")
            return
    
    if not file_path or not os.path.exists(file_path):
        await update.message.reply_text("❌ File not found. Please upload a new document.")
        return
    
    status_msg = await update.message.reply_text("🧪 Testing duplex settings... This may take a moment.")
    
    try:
        settings_to_test = [
            ("duplex", "Standard duplex"),
            ("duplexlong", "Flip on long edge"),
            ("duplexshort", "Flip on short edge"),
        ]
        
        results = []
        network_printer = f"\\\\{PRINTER_IP}\\{PRINTER_NAME}" if PRINTER_IP else PRINTER_NAME
        
        for setting, description in settings_to_test:
            cmd = [
                SUMATRA_PATH,
                "-print-to", network_printer,
                "-print-settings", f"1x,paper=A4,{setting},fit",
                "-silent",
                file_path
            ]
            
            try:
                result = subprocess.run(cmd, timeout=30, capture_output=True, text=True)
                if result.returncode == 0:
                    results.append(f"✅ {description}")
                else:
                    results.append(f"❌ {description}: Failed")
            except Exception as e:
                results.append(f"❌ {description}: Error")
        
        await status_msg.edit_text(
            "**Duplex Test Results:**\n\n" + "\n".join(results) +
            "\n\n*📝 Check your printer to see which setting produced double-sided prints!*",
            parse_mode=ParseMode.MARKDOWN
        )
        
    except Exception as e:
        await status_msg.edit_text(f"❌ Test failed: {str(e)}")

async def check_printer(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Check printer capabilities and current settings"""
    if not await check_thread_access(update, silent=False):
        return
        
    status_msg = await update.message.reply_text("🔍 Checking printer capabilities...")
    
    cmd = [
        "powershell",
        "-Command",
        f"Get-Printer -Name '{PRINTER_NAME}' | Select-Object Name, PrinterStatus, duplexingcapability"
    ]
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
        await status_msg.edit_text(
            f"**Printer Information:**\n```\n{result.stdout}\n```",
            parse_mode=ParseMode.MARKDOWN
        )
    except Exception as e:
        await status_msg.edit_text(f"❌ Error checking printer: {e}")

async def show_debug(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show debug information"""
    if not await check_thread_access(update, silent=False):
        return
        
    info = f"**🔧 Debug Info:**\n\n"
    info += f"✅ Bot running\n"
    info += f"🖨️ Printer: `{PRINTER_NAME}`\n"
    info += f"🌐 Printer IP: `{PRINTER_IP or 'NOT SET'}`\n"
    info += f"📂 Download dir: `{DOWNLOAD_DIR}`\n"
    info += f"📊 Active jobs: `{len(job_manager.jobs)}`\n"
    info += f"🔢 Next job number: `{job_manager.next_job_number}`\n"
    info += f"📅 Today: `{datetime.now().strftime('%Y-%m-%d')}`\n"
    
    await update.message.reply_text(info, parse_mode=ParseMode.MARKDOWN)

async def quick_print_test(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Quick test - prints with default settings"""
    if not await check_thread_access(update, silent=False):
        return
        
    if not update.message or not update.message.reply_to_message:
        await update.message.reply_text(
            "❌ Please reply to a document with /quick_test"
        )
        return
    
    msg = update.message.reply_to_message
    
    file_path = None
    for job_id, job in job_manager.jobs.items():
        if str(msg.message_id) in job_id:
            file_path = job.file_path
            break
    
    if not file_path and msg.document:
        file = msg.document
        file_path = os.path.join(DOWNLOAD_DIR, f"quick_test_{file.file_name}")
        
        status_msg = await update.message.reply_text("📥 Downloading file...")
        try:
            tg_file = await context.bot.get_file(file.file_id)
            await tg_file.download_to_drive(file_path)
            await status_msg.delete()
        except Exception as e:
            await status_msg.edit_text(f"❌ Download failed: {str(e)}")
            return
    
    if not file_path or not os.path.exists(file_path):
        await update.message.reply_text("❌ File not found.")
        return
    
    status_msg = await update.message.reply_text("🖨️ Running quick print test...")
    
    try:
        network_printer = f"\\\\{PRINTER_IP}\\{PRINTER_NAME}" if PRINTER_IP else PRINTER_NAME
        
        cmd = [
            SUMATRA_PATH,
            "-print-to", network_printer,
            "-print-settings", "1x,fit",
            "-silent",
            file_path
        ]
        
        result = subprocess.run(cmd, timeout=30, capture_output=True, text=True)
        
        if result.returncode == 0:
            await status_msg.edit_text(
                "✅ **Quick print test successful!**\n\n1 page should have printed.",
                parse_mode=ParseMode.MARKDOWN
            )
        else:
            await status_msg.edit_text(
                f"❌ **Quick print test failed**\n\nError code: {result.returncode}",
                parse_mode=ParseMode.MARKDOWN
            )
            
    except Exception as e:
        await status_msg.edit_text(f"❌ Test failed: {str(e)}")

# ─────────────────────────────────────────────
# JOB STATUS ENUM
# ─────────────────────────────────────────────
class JobStatus(Enum):
    PENDING = "pending"
    CONFIGURING = "configuring"
    PRINTING = "printing"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"

# ─────────────────────────────────────────────
# JOB DATA CLASS
# ─────────────────────────────────────────────
@dataclass
class PrintJob:
    job_number: int
    file_path: str
    file_name: str
    thread_id: int
    user_id: int
    username: Optional[str]
    status: JobStatus
    created_at: float
    paper_size: Optional[str] = None
    color_mode: Optional[str] = None
    duplex: Optional[str] = None
    copies: Optional[int] = None
    awaiting_custom: bool = False
    completed_at: Optional[float] = None
    error_message: Optional[str] = None
    
    def to_dict(self):
        data = asdict(self)
        data['status'] = self.status.value
        return data
    
    @classmethod
    def from_dict(cls, data):
        data['status'] = JobStatus(data['status'])
        return cls(**data)

# ─────────────────────────────────────────────
# JOB QUEUE MANAGER (WITH DAILY RESET)
# ─────────────────────────────────────────────
class JobQueueManager:
    def __init__(self, state_file: str = "job_state.json"):
        self.jobs: Dict[str, PrintJob] = {}
        self.state_file = state_file
        self._lock = asyncio.Lock()
        self._load_state()
    
    def _load_state(self):
        """Load persistent job state with daily reset"""
        try:
            if os.path.exists(self.state_file):
                with open(self.state_file, 'r') as f:
                    data = json.load(f)
                    
                    today = datetime.now().strftime("%Y-%m-%d")
                    last_date = data.get('last_date', '')
                    
                    if last_date != today:
                        # New day, reset counter
                        self.next_job_number = 1
                        logger.info(f"New day! Resetting job counter to 1")
                    else:
                        self.next_job_number = data.get('next_job_number', 1)
                    
                    logger.info(f"Loaded job state: next number = {self.next_job_number}")
        except Exception as e:
            logger.error(f"Failed to load job state: {e}")
            self.next_job_number = 1
    
    def _save_state(self):
        """Save persistent job state with date"""
        try:
            with open(self.state_file, 'w') as f:
                json.dump({
                    'next_job_number': self.next_job_number,
                    'last_date': datetime.now().strftime("%Y-%m-%d"),
                    'last_saved': datetime.now().isoformat()
                }, f)
        except Exception as e:
            logger.error(f"Failed to save job state: {e}")
    
    async def create_job(self, job_id: str, file_path: str, file_name: str, 
                         thread_id: int, user_id: int, username: Optional[str]) -> PrintJob:
        """Create a new job with unique job number"""
        async with self._lock:
            job_number = self.next_job_number
            self.next_job_number += 1
            self._save_state()
            
            job = PrintJob(
                job_number=job_number,
                file_path=file_path,
                file_name=file_name,
                thread_id=thread_id,
                user_id=user_id,
                username=username,
                status=JobStatus.PENDING,
                created_at=time.time()
            )
            
            self.jobs[job_id] = job
            logger.info(f"Created job #{job_number} for user {username or user_id}")
            return job
    
    async def get_job(self, job_id: str) -> Optional[PrintJob]:
        """Get a job by ID"""
        async with self._lock:
            return self.jobs.get(job_id)
    
    async def update_job(self, job_id: str, **kwargs):
        """Update job attributes"""
        async with self._lock:
            if job_id in self.jobs:
                for key, value in kwargs.items():
                    if hasattr(self.jobs[job_id], key):
                        setattr(self.jobs[job_id], key, value)
                logger.debug(f"Updated job {job_id}: {kwargs}")
    
    async def remove_job(self, job_id: str) -> bool:
        """Remove a job from the queue"""
        async with self._lock:
            if job_id in self.jobs:
                job = self.jobs[job_id]
                try:
                    if os.path.exists(job.file_path):
                        os.remove(job.file_path)
                except Exception as e:
                    logger.error(f"Failed to delete file for job {job_id}: {e}")
                
                del self.jobs[job_id]
                logger.info(f"Removed job {job_id}")
                return True
            return False
    
    async def get_active_jobs(self) -> list:
        """Get all active (non-completed) jobs"""
        async with self._lock:
            return [job for job in self.jobs.values() 
                   if job.status not in [JobStatus.COMPLETED, JobStatus.CANCELLED]]
    
    async def cleanup_old_jobs(self, max_age_hours: int = 24):
        """Remove jobs older than specified hours"""
        async with self._lock:
            current_time = time.time()
            to_remove = []
            
            for job_id, job in self.jobs.items():
                if job.status in [JobStatus.COMPLETED, JobStatus.CANCELLED, JobStatus.FAILED]:
                    age_hours = (current_time - job.completed_at) / 3600 if job.completed_at else 0
                    if age_hours > max_age_hours:
                        to_remove.append(job_id)
            
            for job_id in to_remove:
                await self.remove_job(job_id)
            
            if to_remove:
                logger.info(f"Cleaned up {len(to_remove)} old jobs")

# Initialize job queue manager
job_manager = JobQueueManager()

# ─────────────────────────────────────────────
# PRICE CALCULATION
# ─────────────────────────────────────────────
def get_page_count(file_path):
    try:
        if file_path.lower().endswith(".pdf"):
            return len(PdfReader(file_path).pages)
        return 1
    except Exception as e:
        logger.error(f"Failed to get page count: {e}")
        return 1

def calculate_price(job: PrintJob):
    """Calculate price with limits validation"""
    if not all([job.paper_size, job.color_mode, job.duplex, job.copies]):
        return 0, 0
    
    base = 200 if job.color_mode == "BW" else 300
    
    if job.paper_size == "A3":
        base *= 2
    
    pages = get_page_count(job.file_path)
    
    if pages > MAX_PAGES:
        return None, f"Document has {pages} pages, maximum allowed is {MAX_PAGES}"
    
    total = pages * job.copies * base
    
    if job.duplex == "Double" and job.paper_size == "A4":
        total = int(total * 1.25)
    
    return total, pages

# ─────────────────────────────────────────────
# NETWORK PRINTING WITH LOCAL PRINTER NAME
# ─────────────────────────────────────────────
def send_to_printer(job: PrintJob):
    """
    Send job to printer using SumatraPDF with local printer name
    This uses the printer already installed in Windows
    """
    # Validate SumatraPDF exists
    if not os.path.exists(SUMATRA_PATH):
        return False, f"SumatraPDF not found at: {SUMATRA_PATH}"
    
    # Validate file exists
    if not os.path.exists(job.file_path):
        return False, f"File not found: {job.file_path}"
    
    file_size = os.path.getsize(job.file_path)
    if file_size == 0:
        return False, "File is empty"
    
    # Build print settings
    paper = "A4" if job.paper_size == "A4" else "A3"
    
    settings_parts = [
        f"{job.copies}x",
        f"paper={paper}",
        "fit"  # Fit to page - scales content to fit paper size
    ]
    
    # Add duplex if requested
    if job.duplex == "Double":
        settings_parts.append("duplex")
    
    # Add color mode - CRITICAL for BW printing
    if job.color_mode == "BW":
        settings_parts.append("monochrome")
    # Color is default, no need to specify
    
    settings = ",".join(settings_parts)
    
    # Use the local printer name directly (not network path)
    cmd = [
        SUMATRA_PATH,
        "-print-to", PRINTER_NAME,  # Use the local printer name from .env
        "-print-settings", settings,
        "-silent",
        job.file_path
    ]
    
    logger.info(f"Print command: {' '.join(cmd)}")
    logger.info(f"Settings: {settings}")
    logger.info(f"Paper: {paper}, Color: {job.color_mode}, Duplex: {job.duplex}, Copies: {job.copies}")
    
    try:
        result = subprocess.run(cmd, timeout=PRINT_TIMEOUT, capture_output=True, text=True)
        
        if result.stdout:
            logger.info(f"Print stdout: {result.stdout}")
        if result.stderr:
            logger.warning(f"Print stderr: {result.stderr}")
        
        if result.returncode == 0:
            logger.info(f"Job #{job.job_number} sent to printer successfully")
            return True, "OK"
        else:
            error_msg = f"Printer error (code {result.returncode})"
            if result.stderr:
                error_msg += f": {result.stderr}"
            return False, error_msg
            
    except subprocess.TimeoutExpired:
        logger.error(f"Job #{job.job_number} timed out")
        return False, f"Print job timed out after {PRINT_TIMEOUT} seconds"
    except Exception as e:
        logger.error(f"Print error: {e}")
        return False, f"Print failed: {str(e)}"

# ─────────────────────────────────────────────
# SAFE DOWNLOAD
# ─────────────────────────────────────────────
async def safe_download(file, path, context):
    try:
        tg_file = await context.bot.get_file(file.file_id)
        await tg_file.download_to_drive(path)
        
        if os.path.exists(path) and os.path.getsize(path) > 0:
            logger.info(f"Downloaded file: {path} ({os.path.getsize(path)} bytes)")
            return True
        else:
            logger.error(f"Downloaded file is empty or missing: {path}")
            return False
    except Exception as e:
        logger.error(f"Download error: {e}")
        return False

# ─────────────────────────────────────────────
# UI HELPERS
# ─────────────────────────────────────────────
def cancel_btn(job_id):
    return [InlineKeyboardButton(translate("cancel"), callback_data=f"{job_id}|cancel")]

# ─────────────────────────────────────────────
# FILE HANDLER (SILENT)
# ─────────────────────────────────────────────
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    if not msg:
        return
    
    # Check thread access - SILENT (no error message)
    if not await check_thread_access(update, silent=True):
        return  # Just ignore, don't process
    
    # Check file type
    if not (msg.document or msg.photo):
        return
    
    if msg.document:
        file = msg.document
        name = file.file_name or "file"
    else:
        file = msg.photo[-1]
        name = "photo.jpg"
    
    # Validate extension
    if not any(name.lower().endswith(ext) for ext in ALLOWED_EXT):
        await msg.reply_text(translate("unsupported_file"))
        return
    
    # Generate unique job ID
    job_id = f"{file.file_unique_id}_{msg.message_id}"
    path = os.path.join(DOWNLOAD_DIR, name)
    
    # Download file
    ok = await safe_download(file, path, context)
    if not ok:
        await msg.reply_text(translate("download_failed"))
        return
    
    # Create job in queue
    user = msg.from_user
    job = await job_manager.create_job(
        job_id=job_id,
        file_path=path,
        file_name=name,
        thread_id=msg.message_thread_id,
        user_id=user.id,
        username=user.username or f"{user.first_name} {user.last_name or ''}".strip()
    )
    
    await job_manager.update_job(job_id, status=JobStatus.CONFIGURING)
    
    # Show paper size selection with Khmer
    escaped_name = escape_markdown(name)
    escaped_username = escape_markdown(job.username)
    
    keyboard = [
        [
            InlineKeyboardButton(translate("a4"), callback_data=f"{job_id}|paper|A4"),
            InlineKeyboardButton(translate("a3"), callback_data=f"{job_id}|paper|A3"),
        ],
        cancel_btn(job_id)
    ]
    
    try:
        await msg.reply_text(
            f"{translate('job', job_number=job.job_number)}\n"
            f"{translate('file', file_name=escaped_name)}\n"
            f"{translate('user', username=escaped_username)}\n\n"
            f"{translate('select_paper')}",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode=ParseMode.MARKDOWN
        )
    except Exception as e:
        logger.error(f"Markdown error: {e}")
        await msg.reply_text(
            f"🖨️ Job #{job.job_number}\n📄 {name}\n👤 User: {job.username}\n\nSelect paper size:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )

# ─────────────────────────────────────────────
# CALLBACK HANDLER (SILENT)
# ─────────────────────────────────────────────
async def handle_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    
    # Check thread access - SILENT for callbacks
    if not await check_thread_access(update, silent=True):
        await query.answer()  # Still need to answer callback to prevent "loading" state
        return
    
    await query.answer()
    
    parts = query.data.split("|")
    if len(parts) < 2:
        return
    
    job_id = parts[0]
    step = parts[1]
    
    # Get job from manager
    job = await job_manager.get_job(job_id)
    if not job:
        await query.edit_message_text(translate("expired"))
        return
    
    # Handle cancellation
    if step == "cancel":
        await job_manager.update_job(job_id, status=JobStatus.CANCELLED)
        await job_manager.remove_job(job_id)
        await query.edit_message_text(translate("cancelled", job_number=job.job_number))
        return
    
    if len(parts) < 3:
        return
    
    val = parts[2]
    
    try:
        if step == "paper":
            await job_manager.update_job(job_id, paper_size=val)
            
            keyboard = [
                [
                    InlineKeyboardButton(translate("color"), callback_data=f"{job_id}|color|Color"),
                    InlineKeyboardButton(translate("bw"), callback_data=f"{job_id}|color|BW"),
                ],
                cancel_btn(job_id)
            ]
            await query.edit_message_text(
                f"{translate('select_color')}\n\n"
                f"📄 Size: {job.paper_size}",
                reply_markup=InlineKeyboardMarkup(keyboard),
                parse_mode=ParseMode.MARKDOWN
            )
        
        elif step == "color":
            await job_manager.update_job(job_id, color_mode=val)
            
            keyboard = [
                [
                    InlineKeyboardButton(translate("single"), callback_data=f"{job_id}|duplex|Single"),
                    InlineKeyboardButton(translate("double"), callback_data=f"{job_id}|duplex|Double"),
                ],
                cancel_btn(job_id)
            ]
            await query.edit_message_text(
                f"{translate('select_duplex')}\n\n"
                f"📄 Size: {job.paper_size}\n"
                f"🎨 Mode: {job.color_mode}",
                reply_markup=InlineKeyboardMarkup(keyboard),
                parse_mode=ParseMode.MARKDOWN
            )
        
        elif step == "duplex":
            await job_manager.update_job(job_id, duplex=val)
            
            keyboard = [
                [InlineKeyboardButton("1", callback_data=f"{job_id}|copies|1")],
                [InlineKeyboardButton("2", callback_data=f"{job_id}|copies|2")],
                [InlineKeyboardButton("5", callback_data=f"{job_id}|copies|5")],
                [InlineKeyboardButton("10", callback_data=f"{job_id}|copies|10")],
                [InlineKeyboardButton("25", callback_data=f"{job_id}|copies|25")],
                [InlineKeyboardButton("50", callback_data=f"{job_id}|copies|50")],
                [InlineKeyboardButton("100", callback_data=f"{job_id}|copies|100")],
                [InlineKeyboardButton(translate("custom"), callback_data=f"{job_id}|copies|custom")],
                cancel_btn(job_id)
            ]
            await query.edit_message_text(
                f"{translate('select_copies')}\n\n"
                f"📄 Size: {job.paper_size}\n"
                f"🎨 Mode: {job.color_mode}\n"
                f"📄 Duplex: {job.duplex}\n\n"
                f"*អតិបរមា {MAX_COPIES} ច្បាប់*",
                reply_markup=InlineKeyboardMarkup(keyboard),
                parse_mode=ParseMode.MARKDOWN
            )
        
        elif step == "copies":
            if val == "custom":
                await job_manager.update_job(job_id, awaiting_custom=True)
                await query.edit_message_text(
                    f"{translate('enter_copies')}\n\n"
                    f"*អប្បបរមា: {MIN_COPIES}*\n"
                    f"*អតិបរមា: {MAX_COPIES}*\n\n"
                    f"សូមវាយបញ្ចូលលេខក្នុងឆាត៖",
                    reply_markup=InlineKeyboardMarkup([cancel_btn(job_id)]),
                    parse_mode=ParseMode.MARKDOWN
                )
                return
            
            copies = int(val)
            if copies < MIN_COPIES or copies > MAX_COPIES:
                await query.edit_message_text(
                    f"❌ ចំនួនច្បាប់ត្រូវតែចន្លោះពី {MIN_COPIES} ដល់ {MAX_COPIES}",
                    reply_markup=InlineKeyboardMarkup([cancel_btn(job_id)])
                )
                return
            
            await job_manager.update_job(job_id, copies=copies)
            await finalize_job(query, job_id)
    
    except Exception as e:
        logger.error(f"Error in callback: {e}")
        await query.edit_message_text("❌ កំហុសកើតឡើង។ សូមព្យាយាមម្តងទៀត។")

# ─────────────────────────────────────────────
# FINALIZE JOB
# ─────────────────────────────────────────────
async def finalize_job(message_or_query, job_id):
    """Finalize and print the job"""
    job = await job_manager.get_job(job_id)
    if not job:
        return
    
    # Validate all settings
    if not all([job.paper_size, job.color_mode, job.duplex, job.copies]):
        error_text = "❌ Missing configuration. Please start over."
        if hasattr(message_or_query, 'edit_message_text'):
            await message_or_query.edit_message_text(error_text)
        else:
            await message_or_query.reply_text(error_text)
        return
    
    await job_manager.update_job(job_id, status=JobStatus.PRINTING)
    
    # Calculate price with validation
    price_result, pages = calculate_price(job)
    if price_result is None:
        await job_manager.update_job(job_id, status=JobStatus.FAILED, error_message=pages)
        error_text = f"❌ {pages}"
        if hasattr(message_or_query, 'edit_message_text'):
            await message_or_query.edit_message_text(error_text)
        else:
            await message_or_query.reply_text(error_text)
        return
    
    # Send to printer
    success, msg = send_to_printer(job)
    
    if not success:
        await job_manager.update_job(job_id, status=JobStatus.FAILED, error_message=msg)
        error_text = f"{translate('print_failed')}\n\n{msg}"
        if hasattr(message_or_query, 'edit_message_text'):
            await message_or_query.edit_message_text(error_text)
        else:
            await message_or_query.reply_text(error_text)
        return
    
    # Job completed successfully
    await job_manager.update_job(
        job_id, 
        status=JobStatus.COMPLETED,
        completed_at=time.time()
    )
    
    # Create success message with Khmer
    escaped_filename = escape_markdown(job.file_name)
    escaped_username = escape_markdown(job.username)
    
    # Map color mode to Khmer
    color_display = "ពណ៌" if job.color_mode == "Color" else "ខ្មៅស"
    duplex_display = "សងខាង" if job.duplex == "Double" else "ម្ខាង"
    
    success_text = (
        f"{translate('completed', job_number=job.job_number)}\n\n"
        f"{translate('file', file_name=escaped_filename)}\n"
        f"{translate('pages', pages=pages)}\n"
        f"{translate('copies', copies=job.copies)}\n"
        f"{translate('size', paper_size=job.paper_size)}\n"
        f"🎨 ពណ៌: {color_display}\n"
        f"📄 បោះពុម្ព: {duplex_display}\n"
        f"{translate('price', price=price_result)}\n\n"
        f"{translate('user', username=escaped_username)}\n"
        f"{translate('printed', time=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))}"
    )
    
    try:
        if hasattr(message_or_query, 'edit_message_text'):
            await message_or_query.edit_message_text(
                success_text,
                parse_mode=ParseMode.MARKDOWN
            )
        else:
            await message_or_query.reply_text(
                success_text,
                parse_mode=ParseMode.MARKDOWN
            )
    except Exception as e:
        logger.error(f"Markdown error in finalize: {e}")
        plain_text = (
            f"✅ Job #{job.job_number} Completed!\n\n"
            f"📄 File: {job.file_name}\n"
            f"📄 Pages: {pages}\n"
            f"📄 Copies: {job.copies}\n"
            f"📄 Size: {job.paper_size}\n"
            f"🎨 Mode: {color_display}\n"
            f"📄 Duplex: {duplex_display}\n"
            f"💰 Price: {price_result:,} riels\n\n"
            f"👤 User: {job.username}\n"
            f"🕐 Printed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )
        if hasattr(message_or_query, 'edit_message_text'):
            await message_or_query.edit_message_text(plain_text)
        else:
            await message_or_query.reply_text(plain_text)
    
    # Remove job after successful print
    await job_manager.remove_job(job_id)
    logger.info(f"Job #{job.job_number} completed and cleaned up")

# ─────────────────────────────────────────────
# TEXT HANDLER FOR CUSTOM COPIES (SILENT)
# ─────────────────────────────────────────────
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle custom copies input"""
    text = update.message.text.strip()
    user_id = update.message.from_user.id
    
    # Check thread access - SILENT
    if not await check_thread_access(update, silent=True):
        return
    
    # Look for jobs waiting for custom copies from THIS user
    found_job = None
    for job_id, job in list(job_manager.jobs.items()):
        if job.awaiting_custom and job.user_id == user_id:
            found_job = (job_id, job)
            break
    
    if not found_job:
        # Not a custom copies response, ignore
        return
    
    job_id, job = found_job
    
    # Validate input
    if not text.isdigit():
        await update.message.reply_text(
            f"❌ សូមបញ្ចូលលេខចន្លោះពី {MIN_COPIES} ដល់ {MAX_COPIES}",
            reply_markup=InlineKeyboardMarkup([cancel_btn(job_id)])
        )
        return
    
    copies = int(text)
    if copies < MIN_COPIES or copies > MAX_COPIES:
        await update.message.reply_text(
            f"❌ ចំនួនច្បាប់ត្រូវតែចន្លោះពី {MIN_COPIES} ដល់ {MAX_COPIES}",
            reply_markup=InlineKeyboardMarkup([cancel_btn(job_id)])
        )
        return
    
    # Update job and finalize
    await job_manager.update_job(job_id, copies=copies, awaiting_custom=False)
    
    # Send a temporary "processing" message
    processing_msg = await update.message.reply_text("🖨️ កំពុងដំណើរការ...")
    
    try:
        await finalize_job(update.message, job_id)
        await processing_msg.delete()
    except Exception as e:
        await processing_msg.edit_text(f"❌ កំហុស: {str(e)}")

# ─────────────────────────────────────────────
# PERIODIC CLEANUP
# ─────────────────────────────────────────────
async def cleanup_task(context: ContextTypes.DEFAULT_TYPE):
    """Periodically clean up old jobs"""
    await job_manager.cleanup_old_jobs(max_age_hours=24)
    logger.debug("Ran periodic job cleanup")

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    # Validate required config
    if not BOT_TOKEN:
        raise ValueError("BOT_TOKEN not found in environment variables")
    if not PRINTER_IP:
        logger.warning("PRINTER_IP not set. Network printing will not work!")
    
    # Create the application
    app = Application.builder().token(BOT_TOKEN).build()
    
    # IMPORTANT: Add command handlers FIRST
    # Basic command
    app.add_handler(MessageHandler(filters.Regex(r'^/ping$'), ping))
    
    # Debug commands
    app.add_handler(MessageHandler(filters.Regex(r'^/test_duplex$'), test_duplex))
    app.add_handler(MessageHandler(filters.Regex(r'^/check_printer$'), check_printer))
    app.add_handler(MessageHandler(filters.Regex(r'^/debug$'), show_debug))
    app.add_handler(MessageHandler(filters.Regex(r'^/quick_test$'), quick_print_test))
    
    # Callback handler for button clicks
    app.add_handler(CallbackQueryHandler(handle_callback))
    
    # Text handler for custom copies (non-command text)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    
    # File handler - handle documents and photos (not commands)
    app.add_handler(MessageHandler(
        (filters.Document.ALL | filters.PHOTO) & ~filters.COMMAND, 
        handle_file
    ))
        
    # Add periodic cleanup job (every 6 hours)
    job_queue = app.job_queue
    if job_queue:
        job_queue.run_repeating(cleanup_task, interval=21600, first=10)
    
    logger.info("=" * 50)
    logger.info("Bot started successfully!")
    logger.info(f"Printer: {PRINTER_NAME}")
    logger.info(f"Printer IP: {PRINTER_IP or 'NOT SET - Network printing will not work'}")
    logger.info(f"Network printer path: \\\\{PRINTER_IP}\\{PRINTER_NAME}" if PRINTER_IP else "Network printer path: NOT SET")
    logger.info("Commands available:")
    logger.info("  /ping - Check if bot is responding")
    logger.info("  /debug - Show debug info")
    logger.info("  /check_printer - Check printer capabilities")
    logger.info("  /quick_test - Quick print test (reply to document)")
    logger.info("  /test_duplex - Test duplex settings (reply to document)")
    logger.info("=" * 50)
    app.run_polling()

if __name__ == "__main__":
    main()