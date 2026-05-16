import json
import requests
from pathlib import Path
from datetime import datetime, timezone
import pandas as pd
from openpyxl.utils import get_column_letter
from textual import work
from textual.app import App, ComposeResult
from textual.screen import ModalScreen
from textual.containers import Vertical, Horizontal
from textual.widgets import Header, Footer, DataTable, TabbedContent, TabPane, Input, Button, Label

OUTPUT_DIR = "exophase_json"

class SyncModal(ModalScreen[str]):
    CSS = """
    SyncModal {
        align: center middle;
        background: $background 80%;
    }
    #dialog {
        width: 60;
        height: auto;
        padding: 2;
        background: $surface;
        border: solid $accent;
    }
    Horizontal {
        margin-top: 1;
        height: auto;
        align: right middle;
    }
    Button { margin-left: 1; }
    """

    def compose(self) -> ComposeResult:
        with Vertical(id="dialog"):
            yield Label("Синхронізація з Exophase")
            yield Input(placeholder="Player ID", id="pid")
            with Horizontal():
                yield Button("Скасувати(Cancel)", variant="error", id="cancel")
                yield Button("Синхронізувати(Sync)", variant="success", id="start")

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "start":
            pid = self.query_one("#pid", Input).value.strip()
            if pid:
                self.dismiss(pid)
            else:
                self.app.notify("Введіть Player ID!", severity="warning")
        else:
            self.dismiss(None)


class LocalExophase(App):
    CSS = """
    DataTable { width: 1fr; height: 1fr; }
    """
    BINDINGS = [
        ("q", "quit", "Вихід(Quit)"),
        ("s", "sync", "Синхронізувати(Sync)"),
        ("d", "delete_game", "Видалити(Delete)"),
        ("e", "export", "Експорт(Export) в Excel")
    ]

    def __init__(self):
        super().__init__()
        self.sort_state = {}
        self.games_data = []
        self.latest_json_path = None

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True)
        with TabbedContent():
            with TabPane("PlayStation", id="tp-ps"):
                yield DataTable(id="dt-ps")
            with TabPane("Xbox", id="tp-xbox"):
                yield DataTable(id="dt-xbox")
            with TabPane("Steam", id="tp-steam"):
                yield DataTable(id="dt-steam")
            with TabPane("Other", id="tp-other"):
                yield DataTable(id="dt-other")
        yield Footer()

    def on_mount(self) -> None:
        tables = {
            "dt-ps": ("Game", "Playtime", "Bronze", "Silver", "Gold", "Platinum", "Completion %", "Last Played (UTC)"),
            "dt-xbox": ("Game", "Playtime", "Earned Awards", "Total Awards", "Earned Points", "Completion %", "Last Played (UTC)"),
            "dt-steam": ("Game", "Playtime", "Earned Awards", "Total Awards", "Completion %", "Last Played (UTC)"),
            "dt-other": ("Game", "Playtime", "Platforms", "Completion %", "Last Played (UTC)")
        }

        for dt_id, cols in tables.items():
            dt = self.query_one(f"#{dt_id}", DataTable)
            dt.cursor_type = "row"
            dt.zebra_stripes = True
            dt.add_columns(*cols)
            
        self.call_after_refresh(self.load_data)

    def load_data(self):
        for dt_id in ["dt-ps", "dt-xbox", "dt-steam", "dt-other"]:
            self.query_one(f"#{dt_id}", DataTable).clear()

        base_dir = Path(__file__).resolve().parent
        files = list(base_dir.rglob("all_games_*.json"))
        
        if not files:
            self.games_data = []
            return

        self.latest_json_path = max(files, key=lambda p: p.stat().st_mtime)
        
        try:
            with open(self.latest_json_path, 'r', encoding='utf-8') as f:
                self.games_data = json.load(f)
        except Exception as e:
            self.notify(f"Помилка: {e}", severity="error")
            return

        dt_ps = self.query_one("#dt-ps", DataTable)
        dt_xbox = self.query_one("#dt-xbox", DataTable)
        dt_steam = self.query_one("#dt-steam", DataTable)
        dt_other = self.query_one("#dt-other", DataTable)

        loaded_count = 0

        for game in self.games_data:
            if not isinstance(game, dict):
                continue

            meta = game.get("meta") or {}
            title = meta.get("title", "Unknown")
            
            p_units = game.get("playtimeUnits") or {}
            playtime_h = p_units.get("hours", 0)
            playtime_m = p_units.get("minutes", 0)
            playtime = f"{playtime_h}h {playtime_m}m" if (playtime_h or playtime_m) else str(game.get("playtime", "0h"))
            
            percent = game.get("percent", 0.0)

            ts = game.get("lastplayed_utc", 0)
            if ts and ts > 0:
                lastplayed = datetime.fromtimestamp(ts, timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
            else:
                lastplayed = ""

            platforms_data = meta.get("platforms") or []
            platforms_lower = [str(p.get("name", "")).lower() for p in platforms_data if isinstance(p, dict)]
            platforms_display = ", ".join([str(p.get("name", "")) for p in platforms_data if isinstance(p, dict)])
            
            if any(kw in p for p in platforms_lower for kw in ["playstation", "ps4", "ps5", "ps3", "ps vita"]):
                row = (
                    title, playtime,
                    game.get("earned_bronze", 0) or 0,
                    game.get("earned_silver", 0) or 0,
                    game.get("earned_gold", 0) or 0,
                    game.get("earned_platinum", 0) or 0,
                    percent, lastplayed
                )
                dt_ps.add_row(*row)
            elif any("xbox" in p for p in platforms_lower):
                row = (
                    title, playtime,
                    game.get("earned_awards", 0) or 0,
                    game.get("total_awards", 0) or 0,
                    game.get("earned_points", 0) or 0,
                    percent, lastplayed
                )
                dt_xbox.add_row(*row)
            elif any("steam" in p for p in platforms_lower):
                row = (
                    title, playtime,
                    game.get("earned_awards", 0) or 0,
                    game.get("total_awards", 0) or 0,
                    percent, lastplayed
                )
                dt_steam.add_row(*row)
            else:
                row = (
                    title, playtime, platforms_display,
                    percent, lastplayed
                )
                dt_other.add_row(*row)
                
            loaded_count += 1

        self.notify(f"Завантажено {loaded_count} ігор", severity="information")

    def sort_cell_value(self, value):
        if isinstance(value, (int, float)):
            return value
        if isinstance(value, str):
            if value == "":
                return 0
            if "h " in value and "m" in value:
                try:
                    h, m = value.split("h ")
                    m = m.replace("m", "")
                    return int(h) * 60 + int(m)
                except ValueError:
                    return 0
            try:
                return float(value)
            except ValueError:
                return value.lower()
        return value

    def on_data_table_header_selected(self, event: DataTable.HeaderSelected):
        dt = event.data_table
        col_key = event.column_key
        state_key = f"{dt.id}_{col_key}"
        
        reverse = self.sort_state.get(state_key, False)
        dt.sort(col_key, key=self.sort_cell_value, reverse=reverse)
        self.sort_state[state_key] = not reverse

    def save_current_data(self):
        base_dir = Path(__file__).resolve().parent
        out_dir = base_dir / OUTPUT_DIR
        out_dir.mkdir(exist_ok=True)
        
        if not self.latest_json_path:
            self.latest_json_path = out_dir / "all_games_manual.json"
            
        try:
            with open(self.latest_json_path, "w", encoding="utf-8") as f:
                json.dump(self.games_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.notify(f"Помилка збереження: {e}", severity="error")

    def action_delete_game(self):
        tabs = self.query_one(TabbedContent)
        active_tab_id = tabs.active
        
        if not active_tab_id:
            return

        dt_map = {
            "tp-ps": "#dt-ps",
            "tp-xbox": "#dt-xbox",
            "tp-steam": "#dt-steam",
            "tp-other": "#dt-other"
        }
        
        dt = self.query_one(dt_map.get(active_tab_id, "#dt-other"), DataTable)
        
        try:
            row_key = dt.coordinate_to_cell_key(dt.cursor_coordinate).row_key
            row_values = dt.get_row(row_key)
            target_title = str(row_values[0])
            
            for i, game in enumerate(self.games_data):
                meta = game.get("meta") or {}
                if meta.get("title", "") == target_title:
                    del self.games_data[i]
                    break
                    
            self.save_current_data()
            self.load_data()
            self.notify(f"Видалено: {target_title}", severity="information")
        except Exception:
            self.notify("Немає виділеного рядка для видалення", severity="warning")

    def action_sync(self):
        def check_sync(player_id: str | None):
            if player_id:
                self.notify("Синхронізація...")
                self.fetch_api_data(player_id)

        self.push_screen(SyncModal(), check_sync)

    def action_export(self):
        self.notify("Формування Excel файлу...")
        self.export_excel_data()

    @work(thread=True)
    def fetch_api_data(self, player_id: str):
        base_dir = Path(__file__).resolve().parent
        out_dir = base_dir / OUTPUT_DIR
        out_dir.mkdir(exist_ok=True)

        headers = {
            "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "en-US,en;q=0.9",
            "Referer": f"https://www.exophase.com/"
        }

        all_games = []
        page_num = 1

        while True:
            url = f"https://api.exophase.com/public/player/{player_id}/games?page={page_num}&environment=&sort=1&showHidden=0"
            try:
                resp = requests.get(url, headers=headers, timeout=10)
                resp.raise_for_status()
                data = resp.json()
            except Exception as e:
                self.call_from_thread(self.notify, f"Помилка API: {e}", severity="error")
                break

            if not data.get("success", False):
                break

            games = data.get("games", [])
            if not games:
                break

            all_games.extend(games)
            page_num += 1

        if all_games:
            json_filepath = out_dir / f"all_games_{player_id}.json"
            try:
                with open(json_filepath, "w", encoding="utf-8") as f:
                    json.dump(all_games, f, ensure_ascii=False, indent=4)
                self.call_from_thread(self.notify, f"Успіх: {len(all_games)} ігор", severity="information")
                self.call_from_thread(self.load_data)
            except Exception as e:
                self.call_from_thread(self.notify, f"Помилка збереження: {e}", severity="error")

    @work(thread=True)
    def export_excel_data(self):
        base_dir = Path(__file__).resolve().parent
        out_dir = base_dir / OUTPUT_DIR
        
        files = list(out_dir.rglob("all_games_*.json"))
        if not files:
            self.call_from_thread(self.notify, "Немає JSON файлів для експорту", severity="error")
            return

        latest_file = max(files, key=lambda p: p.stat().st_mtime)
        player_id = latest_file.stem.split("_")[-1]

        try:
            with open(latest_file, 'r', encoding='utf-8') as f:
                all_games = json.load(f)
        except Exception as e:
            self.call_from_thread(self.notify, f"Помилка читання JSON: {e}", severity="error")
            return

        ps_keywords = ["playstation", "ps4", "ps5", "ps3", "ps vita"]
        xbox_keywords = ["xbox"]
        steam_keywords = ["steam"]

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

            lastplayed_ts = game.get("lastplayed_utc", 0)
            lastplayed_str = ""
            if lastplayed_ts and lastplayed_ts > 0:
                try:
                    lastplayed_str = datetime.fromtimestamp(lastplayed_ts, timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    lastplayed_str = ""

            completion_percent = game.get("percent", 0.0)

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

        excel_filepath = out_dir / f"exophase_games_{player_id}.xlsx"

        try:
            with pd.ExcelWriter(excel_filepath, engine="openpyxl") as writer:
                for sheet_name, rows in sheets_data.items():
                    if rows:
                        df = pd.DataFrame(rows)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        worksheet = writer.sheets[sheet_name]
                        for col_idx, col in enumerate(worksheet.columns, 1):
                            max_length = 0
                            column_letter = get_column_letter(col_idx)
                            for cell in col:
                                if cell.value:
                                    max_length = max(max_length, len(str(cell.value)))
                            worksheet.column_dimensions[column_letter].width = max_length + 2

            self.call_from_thread(self.notify, f"Експортовано в {excel_filepath.name}", severity="information")
        except Exception as e:
            self.call_from_thread(self.notify, f"Помилка експорту: {e}", severity="error")

if __name__ == "__main__":
    app = LocalExophase()
    app.run()
