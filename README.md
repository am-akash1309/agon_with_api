# Invoice Assistant (Flask API + Agno)


## 📁 Project Structure

```
.
├── backend/
│   ├── server.py             # Backend server logic
│   └── ...
├── app.py                    # Main assistant logic
├── tools.py                  # Helper functions and utilities
├── requirements.txt          # Python dependencies
└── .env                      # Environment variables (to be created)
```


## 🔧 Setup Instructions

### 1. Clone the Repository

### 2. Install Dependencies

### 3. Create a `.env` File

Create a `.env` file in the root directory with the following content:

```env
GOOGLE_API_KEY=your_google_api_key
TELEGRAM_BOT_TOKEN=your_bot_token
TELEGRAM_CHAT_ID=your_telegram_chat_id
```

> 🧾 **To get your Telegram Chat ID**: Message [@userinfobot](https://t.me/userinfobot) on Telegram and send `/start`.

> 🧾 **Go to Blue Bot**: Message [@Blue1309Bot](https://t.me/Blue1309Bot) on Telegram and click `/start` to start receving messages.

---

## ▶️ Running the Application

### Step 1: Start the Backend Server

```bash
cd backend
python server.py
```

### Step 2: Run the Assistant

In the project root:

```bash
python app.py
```


