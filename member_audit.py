import discord
import pandas as pd
from discord.ext import commands
import asyncio
from typing import Dict, Set
from dotenv import load_dotenv
import os
import sys
from datetime import datetime

# Load environment variables
load_dotenv()
print("Environment variables loaded")

class MemberAudit(commands.Bot):
    def __init__(self):
        print("Initializing bot...")
        intents = discord.Intents.default()
        intents.members = True
        intents.guilds = True
        intents.message_content = True
        intents.presences = True
        
        super().__init__(command_prefix='!', intents=intents)
        self.audit_complete = False
        print("Bot initialized")

    def get_member_display_name(self, member: discord.Member) -> str:
        """Get the member's display name (same as what Sesh sees)"""
        return member.display_name  # This will return nickname if set, otherwise username

    def export_results(self, all_members: Dict, sesh_responses: Dict):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = r"C:\Users\Will\Desktop\Projects\sesh\audit_results_{}.xlsx".format(timestamp)
        
        results_data = []
        
        for display_name, member in all_members.items():
            results_data.append({
                'Display Name': display_name,
                'Sesh Response': sesh_responses.get(display_name, 'No Response')
            })
        
        df = pd.DataFrame(results_data)
        df = df.sort_values('Display Name', ignore_index=True)
        
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Audit Results', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Audit Results']
            
            # Formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4B8BBE',
                'font_color': 'white',
                'border': 1
            })
            
            yes_format = workbook.add_format({
                'bg_color': '#90EE90',  # Light green
                'border': 1
            })
            
            maybe_format = workbook.add_format({
                'bg_color': '#FFD700',  # Gold
                'border': 1
            })
            
            no_format = workbook.add_format({
                'bg_color': '#FFB6C1',  # Light red
                'border': 1
            })
            
            no_response_format = workbook.add_format({
                'bg_color': '#D3D3D3',  # Light gray
                'border': 1
            })
            
            # Apply header format
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Apply conditional formatting based on response
            for row_num in range(1, len(df) + 1):
                response = df.iloc[row_num-1]['Sesh Response']
                worksheet.write(row_num, 0, df.iloc[row_num-1]['Display Name'])
                
                if response == 'Yes':
                    format_to_use = yes_format
                elif response == 'Maybe':
                    format_to_use = maybe_format
                elif response == 'No':
                    format_to_use = no_format
                else:
                    format_to_use = no_response_format
                    
                worksheet.write(row_num, 1, response, format_to_use)
            
            worksheet.set_column('A:A', 25)
            worksheet.set_column('B:B', 15)
            
        print(f"\nResults exported to: {output_path}")

    async def perform_audit(self):
        print("Beginning audit process...")
        
        GUILD_ID = int(os.getenv('GUILD_ID'))
        CSV_PATH = r"C:\Users\Will\Desktop\Projects\sesh\sesh_export.csv"

        print(f"Looking for guild ID: {GUILD_ID}")
        guild = self.get_guild(GUILD_ID)
        
        if not guild:
            print(f"Error: Could not find guild with ID {GUILD_ID}")
            print(f"Available guilds: {[g.name for g in self.guilds]}")
            return

        print(f"Found guild: {guild.name}")
        
        # Read CSV and create response mapping
        print("Reading CSV file...")
        try:
            df = pd.read_csv(CSV_PATH)
            
            # Create a mapping of usernames to their responses
            sesh_responses = {}
            
            # Check each column for responses
            for index, row in df.iterrows():
                if pd.notna(row['Attendees']):
                    sesh_responses[row['Attendees'].strip()] = 'Yes'
                elif pd.notna(row['Maybe']):
                    sesh_responses[row['Maybe'].strip()] = 'Maybe'
                elif pd.notna(row['No']):
                    sesh_responses[row['No'].strip()] = 'No'
            
            print(f"Found {len(sesh_responses)} responses in CSV")
            print("Responses:", sesh_responses)
        except Exception as e:
            print(f"Error reading CSV: {e}")
            return

        # Get guild members using display names
        print("Getting guild members...")
        all_members = {
            self.get_member_display_name(member): member
            for member in guild.members 
            if not member.bot
        }
        print(f"Found {len(all_members)} guild members")

        # Export results to Excel
        self.export_results(all_members, sesh_responses)

        # Print summary
        response_counts = {
            'Yes': sum(1 for r in sesh_responses.values() if r == 'Yes'),
            'Maybe': sum(1 for r in sesh_responses.values() if r == 'Maybe'),
            'No': sum(1 for r in sesh_responses.values() if r == 'No'),
            'No Response': len(all_members) - len(sesh_responses)
        }
        
        print("\n=== Audit Summary ===")
        print(f"Total guild members: {len(all_members)}")
        print(f"Responses breakdown:")
        print(f"- Yes: {response_counts['Yes']}")
        print(f"- Maybe: {response_counts['Maybe']}")
        print(f"- No: {response_counts['No']}")
        print(f"- No Response: {response_counts['No Response']}")

    async def on_ready(self):
        print(f"Bot {self.user.name} is ready!")
        if not self.audit_complete:
            print("Starting audit from on_ready...")
            await self.perform_audit()
            self.audit_complete = True
            print("Audit complete, shutting down...")
            await self.close()

async def main():
    print("Starting main function...")
    TOKEN = os.getenv('DISCORD_TOKEN')
    
    if not TOKEN:
        print("Error: No Discord token found in .env file")
        return
        
    bot = MemberAudit()
    
    try:
        print("Attempting to start bot...")
        await bot.start(TOKEN)
    except discord.LoginFailure:
        print("Error: Failed to login. Check your token.")
    except Exception as e:
        print(f"Error occurred: {type(e).__name__}")
        print(f"Error details: {str(e)}")
    finally:
        print("Bot shutdown complete")

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nScript terminated by user")
    except Exception as e:
        print(f"Critical error: {str(e)}")