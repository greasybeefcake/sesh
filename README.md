# Discord Sesh Response Audit Bot

A Discord bot that compares server members against Sesh bot event responses and generates an Excel report.

## Files Required

1. `member_audit.py` - The main bot script
2. `.env` - Contains your bot token and server ID
3. `sesh_export.csv` - The exported responses from Sesh

## Installation Steps

### 1. Install Required Packages 
bash
pip install discord.py pandas python-dotenv xlsxwriter

### 2. Create Discord Application & Bot
1. Go to [Discord Developer Portal](https://discord.com/developers/applications)
2. Click "New Application"
3. Name your application and click "Create"
4. Go to "Bot" section in left sidebar
5. Click "Add Bot"
6. Click "Reset Token" and copy the new token
7. Enable these Privileged Gateway Intents:
   - PRESENCE INTENT
   - SERVER MEMBERS INTENT
   - MESSAGE CONTENT INTENT

### 3. Invite Bot to Server
1. Go to "OAuth2" > "URL Generator"
2. Select scopes:
   - `bot`
   - `applications.commands`
3. Select bot permissions:
   - Read Messages/View Channels
   - Read Message History
   - View Server Members
4. Copy generated URL
5. Open URL in browser
6. Select your server and authorize

### 4. Get Your Server ID
1. Enable Developer Mode in Discord:
   - User Settings
   - App Settings
   - Advanced
   - Developer Mode
2. Right-click your server name
3. Click "Copy Server ID"

### 5. Create Project Files

#### Create `.env` file:

DISCORD_TOKEN=your_bot_token_here
GUILD_ID=your_server_id_here
CSV_PATH=path/to/your/sesh_export.csv
OUTPUT_DIR=path/to/your/output/directory

#### Create `member_audit.py`:
[Copy the entire script provided separately]

### 6. Export Sesh Responses
1. Go to your Sesh event in Discord
2. Use Sesh's export command
3. Save as `sesh_export.csv` in project directory

## Running the Bot

1. Open terminal in project directory
2. Run:

bash
python member_audit.py

The bot will:
- Connect to Discord
- Read Sesh export CSV
- Compare members and responses
- Generate Excel report
- Auto-shutdown when complete

## Excel Report Format

The report includes:
- Display Name: Server nickname/username
- Sesh Response: Yes/Maybe/No/No Response

Color coding:
- 🟢 Green = "Yes"
- 🟡 Gold = "Maybe"
- 🔴 Red = "No"
- ⚫ Gray = "No Response"

## Project Structure

your-project-directory/
├── member_audit.py
├── .env
└── sesh_export.csv

## Troubleshooting

### Bot Shows Offline
- Verify bot token
- Check intents are enabled
- Verify server permissions

### Missing Members
- Verify SERVER MEMBERS INTENT
- Check bot permissions

### No Responses Found
- Check CSV file path
- Verify CSV format

### Permission Errors
- Check bot server permissions
- Verify file write permissions

## Security Notes

- Never share your bot token
- Add `.env` to `.gitignore`
- Keep bot permissions minimal
- Update token if compromised

## CSV Format
The script expects Sesh's export format:

csv
Attendees,Maybe,No
username1,,
username2,,
,username3,
,,username4

## Common Issues

1. **ModuleNotFoundError**
   - Run: `pip install [missing_package]`

2. **Permission Denied**
   - Check file permissions
   - Run terminal as administrator

3. **Invalid Token**
   - Reset token in Developer Portal
   - Update `.env` file

4. **Bot Not Responding**
   - Check internet connection
   - Verify bot is online
   - Check Discord status

5. **Name Mismatches**
   - Verify display names match exactly
   - Check for special characters

## Best Practices

1. **Regular Updates**
   - Keep dependencies updated
   - Check for Discord API changes

2. **File Management**
   - Backup `.env` securely
   - Clear old exports regularly
   - Use descriptive filenames

3. **Error Handling**
   - Check console output
   - Verify CSV format
   - Monitor bot status

4. **Security**
   - Rotate bot token periodically
   - Limit bot permissions
   - Secure environment variables

## Support Resources

- [Discord Developer Portal](https://discord.com/developers/docs)
- [discord.py Documentation](https://discordpy.readthedocs.io/)
- [Python dotenv Documentation](https://pypi.org/project/python-dotenv/)
- [Pandas Documentation](https://pandas.pydata.org/docs/)

## Quick Reference

### Environment Variables
DISCORD_TOKEN=your_bot_token_here
GUILD_ID=your_server_id_here
CSV_PATH=path/to/your/sesh_export.csv
OUTPUT_DIR=path/to/your/output/directory

### Required Permissions
- Read Messages/View Channels
- Read Message History
- View Server Members

### Required Intents
- Presence
- Server Members
- Message Content

### Command Line

## Maintenance

### Regular Checks
1. Verify bot permissions
2. Check token validity
3. Update dependencies
4. Test CSV exports
5. Monitor Discord API changes

### Updating
1. Backup `.env`
2. Update packages
3. Test functionality
4. Check permissions
5. Verify exports

## Notes
- Bot auto-shutdowns after report
- Display names are case-sensitive
- Keep backups of important data
- Monitor Discord API changes