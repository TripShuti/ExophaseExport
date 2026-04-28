import asyncio
import json
import os
import re
from datetime import datetime
import pandas as pd
from playwright.async_api import async_playwright
from openpyxl.utils import get_column_letter

OUTPUT_DIR = "exophase_json"

def auto_adjust_column_widths(writer, sheet_name):
    worksheet = writer.sheets[sheet_name]
    for col_idx, col in enumerate(worksheet.columns, 1):
        max_length = 0
        column_letter = get_column_letter(col_idx)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        adjusted_width = max_length + 2
        worksheet.column_dimensions[column_letter].width = adjusted_width

async def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    player_id = None
    all_games = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        player_ids = set()

        def response_handler(response):
            url = response.url
            if "/invalidations" in url and "/public/player/" in url:
                m = re.search(r'/public/player/(\d+)/invalidations', url)
                if m:
                    pid = m.group(1)
                    player_ids.add(pid)
                    print(f"Found player ID: {pid}")

        page.on("response", response_handler)

        print("Opening Exophase login page...")
        await page.goto("https://www.exophase.com/login")

        print("Please log in in the browser and go to your profile page, e.g. exophase.com/user/*username*/ . After, press Enter here...")
        input()

        await asyncio.sleep(5)

        if player_ids:
            player_id = player_ids.pop()
            print(f"Automatically found player ID: {player_id}")
        else:
            player_id = input("Could not find player ID automatically. Please enter player ID manually: ").strip()

        if not player_id or not player_id.isdigit():
            print("Invalid player ID. Exiting.")
            await browser.close()
            return

        print("Starting to download games...")

        page_num = 1
        while True:
            url = f"https://api.exophase.com/public/player/{player_id}/games?page={page_num}&environment=&sort=1&showHidden=0"
            print(f"Loading page {page_num}...")
            try:
                await page.goto(url)
                pre_handle = await page.query_selector("pre")
                if not pre_handle:
                    print(f"No JSON found in <pre> tag on page {page_num}.")
                    break
                json_text = await pre_handle.inner_text()
                data = json.loads(json_text)
            except Exception as e:
                print(f"JSON parsing error on page {page_num}: {e}")
                break

            if not data.get("success", False):
                print("success=False, stopping.")
                break

            games = data.get("games", [])
            if not games:
                print("No games found, stopping.")
                break

            all_games.extend(games)
            page_num += 1

        if not all_games:
            print("No games found.")
            await browser.close()
            return

        json_filepath = os.path.join(OUTPUT_DIR, f"all_games_{player_id}.json")
        with open(json_filepath, "w", encoding="utf-8") as f:
            json.dump(all_games, f, ensure_ascii=False, indent=4)
        print(f"Saved {len(all_games)} games to {json_filepath}")

        # Keywords for platform detection
        ps_keywords = ["playstation", "ps4", "ps5", "ps3", "ps vita"]
        xbox_keywords = ["xbox"]
        steam_keywords = ["steam"]

        # Prepare sheets data
        sheets_data = {
            "PlayStation": [],
            "Xbox": [],
            "Steam": [],
            "Other": []
        }

        for game in all_games:
            title = game.get("meta", {}).get("title", "")
            playtime_hours = game.get("playtimeUnits", {}).get("hours", 0)
            playtime_minutes = game.get("playtimeUnits", {}).get("minutes", 0)
            playtime_str = f"{playtime_hours}h {playtime_minutes}m" if (playtime_hours or playtime_minutes) else game.get("playtime", "")

            platform_objs = game.get("meta", {}).get("platforms", [])
            platform_names = [p.get("name", "") for p in platform_objs]
            platform_names_lower = [p.lower() for p in platform_names]

            # Last played as readable date
            lastplayed_ts = game.get("lastplayed_utc", 0)
            lastplayed_str = ""
            if lastplayed_ts and lastplayed_ts > 0:
                try:
                    from datetime import timezone
                    lastplayed_str = datetime.fromtimestamp(lastplayed_ts, timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    lastplayed_str = ""

            completion_percent = game.get("percent", 0.0)

            # Determine sheet based on platform keywords
            if any(any(keyword in p for keyword in ps_keywords) for p in platform_names_lower):
                sheet_name = "PlayStation"
                row = {
                    "Game": title,
                    "Playtime": playtime_str,
                    "Bronze": game.get("earned_bronze", 0) or 0,
                    "Silver": game.get("earned_silver", 0) or 0,
                    "Gold": game.get("earned_gold", 0) or 0,
                    "Platinum": game.get("earned_platinum", 0) or 0,
                    "Completion %": completion_percent,
                    "Last Played (UTC)": lastplayed_str,
                }
            elif any(any(keyword in p for keyword in xbox_keywords) for p in platform_names_lower):
                sheet_name = "Xbox"
                row = {
                    "Game": title,
                    "Playtime": playtime_str,
                    "Earned Awards": game.get("earned_awards", 0) or 0,
                    "Total Awards": game.get("total_awards", 0) or 0,
                    "Earned Points": game.get("earned_points", 0) or 0,
                    "Completion %": completion_percent,
                    "Last Played (UTC)": lastplayed_str,
                }
            elif any(any(keyword in p for keyword in steam_keywords) for p in platform_names_lower):
                sheet_name = "Steam"
                row = {
                    "Game": title,
                    "Playtime": playtime_str,
                    "Earned Awards": game.get("earned_awards", 0) or 0,
                    "Total Awards": game.get("total_awards", 0) or 0,
                    "Completion %": completion_percent,
                    "Last Played (UTC)": lastplayed_str,
                }
            else:
                sheet_name = "Other"
                row = {
                    "Game": title,
                    "Playtime": playtime_str,
                    "Platforms": ", ".join(platform_names),
                    "Completion %": completion_percent,
                    "Last Played (UTC)": lastplayed_str,
                }

            sheets_data[sheet_name].append(row)

        excel_filepath = os.path.join(OUTPUT_DIR, f"exophase_games_{player_id}.xlsx")

        with pd.ExcelWriter(excel_filepath, engine="openpyxl") as writer:
            for sheet_name, rows in sheets_data.items():
                if rows:
                    df = pd.DataFrame(rows)
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    auto_adjust_column_widths(writer, sheet_name)

        print(f"Data successfully saved to Excel: {excel_filepath}")

        input("Press Enter to close the browser and exit...")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())

