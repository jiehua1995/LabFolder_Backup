#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Author: Jie Hua
# Lab: Imhof Group, BMC, LMU Munich
# Project: LabFolder Complete Backup Tool V4
# Date: 2025-10-26
#
# Description:
# This script allows users to download and locally back up data from a LabFolder ELN instance
# hosted by the Medical Faculty, LMU Munich. It supports the backup of text entries, attached files,
# images, sketches, tables, and well plates. The tool also generates RO-Crate-compliant ELN archives
# and corresponding HTML summary reports for easy offline viewing.
#
# Compatibility:
# Tested on Windows 11 with Python 3.12, optimized for a 1920×1080 screen resolution.
#
# Disclaimer:
# This tool is provided “as is” without any warranty of any kind, express or implied.
# The author, lab, and LMU Munich are not responsible for any data loss, system errors, or
# other damages arising from the use of this software. Users should verify all outputs
# and use the tool at their own risk.

"""
LabFolder Complete Backup Tool V4
"""
import os, sys, json, requests, shutil, zipfile, hashlib, time
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from tqdm import tqdm
from html import escape
import re
import urllib3

# Disable SSL warnings for self-signed certificates
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Optional dependencies
try:
    import pandas as pd
    import openpyxl
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

# Selenium dependencies
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.action_chains import ActionChains
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False
    print("⚠️ Selenium not installed - XLSX download and table/wellplate data will be unavailable. Please make sure to install it.")
    print("💡 Install: pip install selenium")

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

load_dotenv()
requests.packages.urllib3.disable_warnings()

class LabFolderBackup:
    def __init__(self):
        self.base = os.getenv('LABFOLDER_URL').rstrip('/')
        self.user = os.getenv('LABFOLDER_USERNAME')
        self.pwd = os.getenv('LABFOLDER_PASSWORD')
        output_dir = os.getenv('DOWNLOAD_DIR', 'labfolder_backup')
        self.out = Path(output_dir)
        self.session = requests.Session()
        self.session.verify = False
        self.auth_token = None
        
        self.stats = {
            'projects': 0,
            'entries': 0,
            'text_elements': 0,
            'images': 0,
            'files': 0,
            'tables': 0,
            'data_elements': 0,
            'sketches': 0,
            'well_plates': 0,
            'xlsx_downloaded': 0,
            'xlsx_failed': 0,
            'errors': []
        }
        
        self.captured_styles = None
        self.captured_scripts = None
        # number of pixels to nudge up after scrollIntoView to reveal floating buttons
        try:
            self.scroll_nudge = int(os.getenv('SCROLL_NUDGE_PIXELS', '120'))
        except Exception:
            self.scroll_nudge = 120
    
    def login(self):
        """Authenticate"""
        print("\n[1/8] 🔐 Logging in...")
        self.session.get(f"{self.base}/", verify=False)
        csrf = self.session.cookies.get('csrfToken')
        
        r = self.session.post(
            f"{self.base}/access/login",
            data={'principal': self.user, 'password': self.pwd, 'csrfToken': csrf},
            verify=False
        )
        
        if r.status_code == 200:
            print(" ✓ Logged in successfully")
            return True
        print(f" ✗ Login failed: {r.status_code}")
        return False
    
    def capture_auth_token(self):
        """Capture Authorization Token using Selenium"""
        if not SELENIUM_AVAILABLE:
            print(" ⚠️  Selenium not installed - cannot capture token")
            return False
        print("-"*60)
        print(" 🔄 Capturing Authorization Token...")
        print("\n" + "▀"*60)
        print("🖱️  MOUSE CONTROL WARNING - STEP 2/2")
        print("⚠️  A Chrome browser window will open shortly!")
        print("👉 Please move your mouse to:")
        print("   • Another monitor (if available)")
        print("   • Outside the browser window")
        print("   • The browser's title bar (safe zone)")
        print("❌ DO NOT move mouse inside the browser content area!")
        print("⏱️  Automation will take ~30 seconds")
        print("-"*60 + "\n")

        chrome_options = Options()
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--window-size=920,1080')
        chrome_options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})
        
        try:
            driver = webdriver.Chrome(options=chrome_options)
            # driver.maximize_window()
            driver.set_page_load_timeout(60)
            
            driver.get(f"{self.base}/access/login")
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.NAME, 'principal'))
            )
            
            driver.find_element(By.NAME, 'principal').send_keys(self.user)
            driver.find_element(By.NAME, 'password').send_keys(self.pwd)
            driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]').click()
            
            # Sleep for 5 seconds to allow page to load
            time.sleep(5)
            
            # Find an entry with table or wellplate
            entry_with_elements = None
            for entry in self.entries:
                elements = entry.get('elements', [])
                if isinstance(elements, list):
                    for elem_list in elements:
                        if isinstance(elem_list, list):
                            for elem in elem_list:
                                if isinstance(elem, dict) and elem.get('type') in [4, 7]:
                                    entry_with_elements = entry.get('id')
                                    break
                        if entry_with_elements:
                            break
                if entry_with_elements:
                    break
            
            if not entry_with_elements:
                entry_with_elements = self.entries[0].get('id') if self.entries else 151709
            
            driver.get(f"{self.base}/notebook#/entries/{entry_with_elements}")
            time.sleep(5)
            
            from selenium.webdriver.common.keys import Keys
            from selenium.webdriver.common.action_chains import ActionChains

            # Prefer scrolling the internal entries container if present (has the scrollbar)
            try:
                scroll_container = driver.find_element(By.ID, "eln_project_content")
                total_height = driver.execute_script("return arguments[0].scrollHeight", scroll_container)
                container_height = driver.execute_script("return arguments[0].clientHeight", scroll_container)

                if total_height > container_height:
                    # Smoothly scroll from top to bottom to trigger lazy network requests
                    step_size = 200
                    scroll_distance = total_height - container_height
                    total_steps = int(scroll_distance / step_size) + 1
                    for step in range(total_steps):
                        current_position = min((step + 1) * step_size, scroll_distance)
                        driver.execute_script("arguments[0].scrollTop = arguments[1]", scroll_container, current_position)
                        time.sleep(0.25)

                    # final ensure at bottom
                    driver.execute_script("arguments[0].scrollTop = arguments[1]", scroll_container, scroll_distance)
                    time.sleep(1)
                else:
                    # Fallback to body scrolling
                    body = driver.find_element(By.TAG_NAME, "body")
                    actions = ActionChains(driver)
                    actions.move_to_element(body).click().perform()
                    time.sleep(1)
                    for i in range(10):
                        body.send_keys(Keys.PAGE_DOWN)
                        time.sleep(0.3)
                    body.send_keys(Keys.END)
                    time.sleep(2)
            except Exception:
                # If container not found or any error, fallback to basic page scrolling
                body = driver.find_element(By.TAG_NAME, "body")
                actions = ActionChains(driver)
                actions.move_to_element(body).click().perform()
                time.sleep(1)
                for i in range(10):
                    body.send_keys(Keys.PAGE_DOWN)
                    time.sleep(0.3)
                body.send_keys(Keys.END)
                time.sleep(2)
            
            logs = driver.get_log('performance')
            
            for log in logs:
                try:
                    import json as json_lib
                    message = json_lib.loads(log['message'])
                    method = message['message']['method']
                    
                    if method == 'Network.requestWillBeSent':
                        params = message['message']['params']
                        request = params.get('request', {})
                        url = request.get('url', '')
                        headers = request.get('headers', {})
                        
                        if '/api/v2/elements/' in url and 'Authorization' in headers:
                            auth_header = headers['Authorization']
                            if auth_header.startswith('Token '):
                                self.auth_token = auth_header[6:]
                                print(f"    ✓ Captured token: {self.auth_token[:20]}...")
                                driver.quit()
                                return True
                except:
                    pass
            
            driver.quit()
            print(" ⚠️  Could not capture token from network logs")
            return False
            
        except Exception as e:
            print(f" ⚠️  Failed to capture token: {e}")
            try:
                driver.quit()
            except:
                pass
            return False
    
    # It is a bit weird that only table and wellplate support API fetching, so we add these two methods here.

    def fetch_table_content_api(self, stable_id):
        """Fetch table content using API"""
        if not hasattr(self, 'auth_token') or not self.auth_token:
            return None
        
        try:
            base_domain = self.base.rsplit('/eln', 1)[0]
            url = f"{base_domain}/api/v2/elements/table/{stable_id}"
            
            headers = {
                'Authorization': f'Token {self.auth_token}',
                'Accept': 'application/json',
                'Referer': f'{self.base}/notebook',
            }
            
            r = self.session.get(url, headers=headers, verify=False, timeout=30)
            
            if r.status_code == 200:
                data = r.json()
                if 'content' in data:
                    return data['content']
            return None
            
        except Exception as e:
            return None
    
    def fetch_wellplate_content_api(self, stable_id):
        """Fetch wellplate content using API"""
        if not hasattr(self, 'auth_token') or not self.auth_token:
            return None
        
        try:
            base_domain = self.base.rsplit('/eln', 1)[0]
            url = f"{base_domain}/api/v2/elements/well-plate/{stable_id}"
            
            headers = {
                'Authorization': f'Token {self.auth_token}',
                'Accept': 'application/json',
                'Referer': f'{self.base}/notebook',
            }
            
            r = self.session.get(url, headers=headers, verify=False, timeout=30)
            
            if r.status_code == 200:
                data = r.json()
                if 'content' in data:
                    return data['content']
            return None
            
        except Exception as e:
            return None
    
    def fetch_data(self):
        """Get all projects and entries"""
        print("\n[2/8] 📥 Fetching data from LabFolder...")
        print("    🔄 Querying projects API...")
        
        proj_resp = self.session.get(f"{self.base}/projects", verify=False).json()
        self.projects = proj_resp.get('dataObj', [])
        self.stats['projects'] = len(self.projects)
        
        print("    🔄 Querying entries API...")
        entry_resp = self.session.get(
            f"{self.base}/filter/list?skipEntries=false",
            verify=False
        ).json()
        self.entries = entry_resp.get('dataObj', {}).get('entries', [])
        self.stats['entries'] = len(self.entries)
        
        print(f"    ✅ Found {self.stats['projects']} projects")
        print(f"    ✅ Found {self.stats['entries']} entries")
        print(f"    📊 Preparing to download all associated content...")

        return self.projects, self.entries
    
    # Images used a different API endpoint
    def download_image(self, block_id, image_hash, save_path):
        """Download image"""
        try:
            url = f"{self.base}/workspace/block/image/{block_id}/{image_hash}"
            r = self.session.get(url, verify=False, timeout=30)
            
            if r.status_code == 200 and len(r.content) > 100:
                with open(save_path, 'wb') as f:
                    f.write(r.content)
                return True
        except Exception as e:
            self.stats['errors'].append(f"Image download failed: {e}")
        return False
    
    def download_file(self, entry_id, element_number, save_path):
        """Download file"""
        try:
            url = f"{self.base}/workspace/block/file/{entry_id}/{element_number}/file"
            r = self.session.get(url, verify=False, stream=True, timeout=30)
            
            if r.status_code == 200:
                with open(save_path, 'wb') as f:
                    for chunk in r.iter_content(8192):
                        f.write(chunk)
                return True
        except Exception as e:
            self.stats['errors'].append(f"File download failed: {e}")
        return False
    
    def process_entries(self):
        """Process all entries and download content"""
        print(f"\n[3/8] 📦 Processing {len(self.entries)} entries...")
        
        for entry in tqdm(self.entries, desc="    Downloading"):
            entry_id = entry['id']
            block_id = entry.get('blockId')
            title = entry.get('title', f'Entry_{entry_id}')
            proj = entry.get('projectName', 'Unknown')
            
            safe_proj = "".join(c for c in proj if c.isalnum() or c in (' ','-','_')).strip() or 'Unknown'
            safe_title = "".join(c for c in title if c.isalnum() or c in (' ','-','_')).strip() or f'Entry_{entry_id}'
            
            entry_dir = self.out / safe_proj / f"{entry_id}_{safe_title}"
            entry_dir.mkdir(parents=True, exist_ok=True)
            
            with open(entry_dir / 'entry_data.json', 'w', encoding='utf-8') as f:
                json.dump(entry, f, indent=2, ensure_ascii=False)
            
            for row_idx, row in enumerate(entry.get('elements', [])):
                for elem in row:
                    elem_type = elem.get('type')
                    elem_id = elem.get('elementId')
                    elem_num = elem.get('elementNumber')
                    meta = elem.get('metaData', {})
                    
                    if elem_type == 1:
                        content = elem.get('content', '')
                        if content:
                            text_file = entry_dir / f'text_{elem_id}.html'
                            with open(text_file, 'w', encoding='utf-8') as f:
                                f.write(content)
                            self.stats['text_elements'] += 1
                    
                    elif elem_type == 2 and block_id:
                        fname = meta.get('fileName', f'image_{elem_id}')
                        elem_title = elem.get('title', '')
                        if elem_title and '.' in elem_title:
                            fname = elem_title
                        
                        fname = "".join(c if c.isalnum() or c in (' ','-','_','.') else '_' for c in fname).strip()
                        if not fname or fname == '.':
                            fname = f'image_{elem_id}'
                        
                        if '.' not in fname:
                            fname += '.png'
                        
                        for hash_key in ['imageOriginalHash', 'imageHash']:
                            img_hash = meta.get(hash_key)
                            if img_hash:
                                save_path = entry_dir / fname
                                if self.download_image(block_id, img_hash, save_path):
                                    self.stats['images'] += 1
                                    break
                        
                        layer_hash = meta.get('layerHash')
                        if layer_hash:
                            layer_fname = fname.replace('.', '_annotated.')
                            layer_path = entry_dir / layer_fname
                            if self.download_image(block_id, layer_hash, layer_path):
                                pass
                    
                    elif elem_type == 3 and elem_num:
                        fname = meta.get('fileName', f'file_{elem_id}')
                        save_path = entry_dir / fname
                        
                        if self.download_file(entry_id, elem_num, save_path):
                            self.stats['files'] += 1
                    
                    elif elem_type == 4:
                        table_title = elem.get('title', 'Table')
                        stable_id = elem.get('stableId')
                        
                        table_content = {}
                        if stable_id:
                            table_content = self.fetch_table_content_api(stable_id) or {}
                        
                        table_data = {
                            'element_id': elem_id,
                            'title': table_title,
                            'type': 'table',
                            'content': table_content,
                            'variables': elem.get('variables', []),
                            'metadata': meta,
                            'stableId': stable_id
                        }
                        
                        table_file = entry_dir / f'table_{elem_id}.json'
                        with open(table_file, 'w', encoding='utf-8') as f:
                            json.dump(table_data, f, indent=2, ensure_ascii=False)
                        
                        self.stats['tables'] += 1
                    
                    elif elem_type == 5:
                        variables_str = elem.get('variables', '')
                        
                        data_info = {
                            'element_id': elem_id,
                            'type': 'data_element',
                            'variables': variables_str,
                            'content': elem.get('content', ''),
                            'metadata': meta
                        }
                        data_file = entry_dir / f'data_{elem_id}.json'
                        with open(data_file, 'w', encoding='utf-8') as f:
                            json.dump(data_info, f, indent=2, ensure_ascii=False)
                        
                        self.stats['data_elements'] += 1
                    
                    elif elem_type == 6:
                        sketch_meta = {
                            'element_id': elem_id,
                            'type': 'sketch',
                            'title': elem.get('title', 'Sketch'),
                            'metadata': meta
                        }
                        sketch_file = entry_dir / f'sketch_{elem_id}_info.json'
                        with open(sketch_file, 'w', encoding='utf-8') as f:
                            json.dump(sketch_meta, f, indent=2, ensure_ascii=False)
                        
                        self.stats['sketches'] += 1
                    
                    elif elem_type == 7:
                        well_plate_title = elem.get('title', 'Well Plate')
                        stable_id = elem.get('stableId')
                        
                        well_plate_content = {}
                        if stable_id:
                            well_plate_content = self.fetch_wellplate_content_api(stable_id) or {}
                        
                        variables_str = elem.get('variables', '')
                        layers_info = []
                        legend_info = []
                        
                        try:
                            if variables_str:
                                import json as json_lib
                                wp_data = json_lib.loads(variables_str) if isinstance(variables_str, str) else variables_str
                                layers_info = wp_data.get('layers', [])
                                legend_info = wp_data.get('legend', [])
                        except:
                            pass
                        
                        wellplate_data = {
                            'element_id': elem_id,
                            'title': well_plate_title,
                            'type': 'wellplate',
                            'content': well_plate_content,
                            'variables': variables_str,
                            'layers': layers_info,
                            'legend': legend_info,
                            'metadata': meta.get('wellPlateMetaData', {}) if 'wellPlateMetaData' in meta else meta,
                            'stableId': stable_id
                        }
                        
                        wellplate_file = entry_dir / f'wellplate_{elem_id}.json'
                        with open(wellplate_file, 'w', encoding='utf-8') as f:
                            json.dump(wellplate_data, f, indent=2, ensure_ascii=False)
                        
                        self.stats['well_plates'] += 1
        
        print("\n    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        print("    📊 Content Download Summary:")
        print("    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        print(f"    📝 Text blocks:      {self.stats['text_elements']:>6} (HTML components)")
        print(f"    🖼️  Images:           {self.stats['images']:>6} (original)")
        print(f"    📎 Files:            {self.stats['files']:>6} (attachments)")
        print(f"    📊 Tables:           {self.stats['tables']:>6} (JSON + XLSX coming)")
        print(f"    🧪 Well Plates:      {self.stats['well_plates']:>6} (JSON + XLSX coming)")
        print(f"    📈 Data elements:    {self.stats['data_elements']:>6}")
        print(f"    ✏️  Sketches:         {self.stats['sketches']:>6}")
        print("    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
        print(f"    💾 All data saved to: {self.out.name}/")
        print("    ✅ Step 3/8 completed successfully!")

    def download_xlsx_files(self):
        """Download XLSX files for tables and wellplates using Selenium"""
        if not SELENIUM_AVAILABLE:
            print("\n[4/8] ⚠️  Selenium not installed - skipping XLSX download")
            return
        
        print("\n[4/8] 📥 Downloading XLSX files for tables and wellplates...")
        
        # Build a list of elements to download
        elements_to_download = []
        
        for entry_dir in self.out.rglob('*/*/'):
            if entry_dir.name == 'eln_export':
                continue

            # Read entry_data.json to get blockId
            entry_data_file = entry_dir / 'entry_data.json'
            if not entry_data_file.exists():
                continue
            
            with open(entry_data_file, 'r', encoding='utf-8') as f:
                entry_data = json.load(f)
            
            block_id = entry_data.get('blockId')
            if not block_id:
                continue
            
            # Check tables
            for table_file in entry_dir.glob('table_*.json'):
                element_id = int(table_file.stem.split('_')[1])
                # Let the server decide the filename
                elements_to_download.append({
                    'type': 'table',
                    'element_id': element_id,
                    'block_id': block_id,
                    'entry_dir': entry_dir
                })
            
            # Check wellplates
            for wellplate_file in entry_dir.glob('wellplate_*.json'):
                element_id = int(wellplate_file.stem.split('_')[1])
                # Let the server decide the filename
                elements_to_download.append({
                    'type': 'wellplate',
                    'element_id': element_id,
                    'block_id': block_id,
                    'entry_dir': entry_dir
                })
        
        if not elements_to_download:
            print("    ✓ All XLSX files already exist (no download needed)")
            return
        
        
        print(f"    📊 Found {len(elements_to_download)} XLSX files to download")
        print(f"    📐 SCROLL_NUDGE_PIXELS: {self.scroll_nudge}px")
        
        print("-"*60)
        print("\n" + "▀"*60)
        print("🖱️  MOUSE CONTROL WARNING")
        print("⚠️  Chrome browser will open automatically!")
        print("👉 Move your mouse away from browser window:")
        print("   • To another monitor (if available)")
        print("   • Outside the browser boundaries")
        print("   • To the browser title bar (safe zone)")
        print("❌ DO NOT interact with the browser!")
        print(f"⏱️  Estimated time: ~{len(elements_to_download)*30} seconds")
        print("-"*60)
        print("🚀 Starting in 3 seconds...\n")
        time.sleep(3)
        
        # Start Selenium WebDriver
        temp_dir = self.out / '.temp_xlsx_downloads'
        temp_dir.mkdir(exist_ok=True)
        
        chrome_options = Options()
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        
        prefs = {
            "download.default_directory": str(temp_dir.absolute()),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": False,
            # Key settings to allow multiple downloads
            "profile.default_content_setting_values.automatic_downloads": 1,
            "profile.content_settings.exceptions.automatic_downloads.*.setting": 1
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        driver = webdriver.Chrome(options=chrome_options)
        
        try:
            # Login
            driver.get(f"{self.base}/access/login")
            time.sleep(2)
            
            driver.find_element(By.NAME, 'principal').send_keys(self.user)
            driver.find_element(By.NAME, 'password').send_keys(self.pwd)
            driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]').click()
            time.sleep(5)

            # Navigate to notebook page
            driver.get(f"{self.base}/notebook")
            time.sleep(5)

            # Scroll to load all entries
            print("    📜 Scrolling to load all entries...")
            scroll_container = driver.find_element(By.ID, "eln_project_content")
            
            total_height = driver.execute_script("return arguments[0].scrollHeight", scroll_container)
            container_height = driver.execute_script("return arguments[0].clientHeight", scroll_container)
            
            if total_height > container_height:
                scroll_distance = total_height - container_height
                step_size = 200
                total_steps = int(scroll_distance / step_size) + 1
                
                print(f"    ⏬ Smooth scrolling in {total_steps} steps...")
                
                for step in range(total_steps):
                    current_position = min((step + 1) * step_size, scroll_distance)
                    driver.execute_script(
                        "arguments[0].scrollTop = arguments[1]", 
                        scroll_container,
                        current_position
                    )
                    time.sleep(0.3)
                
                print("    ⏳ Waiting 10 seconds for lazy loading...")
                time.sleep(10)
                
                driver.execute_script("arguments[0].scrollTop = 0", scroll_container)
                time.sleep(2)
            
            print("    ✓ All entries loaded")

            # Ensure all entries are expanded
            print("    📂 Expanding all entries if needed...")
            try:
                expand_buttons = driver.find_elements(By.CSS_SELECTOR, 'button.expand-entry-button, .entry-expand-button')
                for btn in expand_buttons:
                    try:
                        if btn.is_displayed():
                            btn.click()
                            time.sleep(0.2)
                    except:
                        pass
            except:
                pass
            
            time.sleep(2)
            print("    ✓ Entries expanded")
            
            # Download each XLSX
            for element_info in tqdm(elements_to_download, desc="    Downloading XLSX"):
                try:
                    elem_type = element_info['type']
                    elem_id = element_info['element_id']
                    block_id = element_info['block_id']
                    entry_dir = element_info['entry_dir']

                    print(f"\n      🔍 Processing {elem_type} {elem_id} (block {block_id})...")

                    # Clear temporary directory (before each download)
                    for f in temp_dir.glob('*'):
                        try:
                            f.unlink()
                        except:
                            pass
                    
                    time.sleep(0.3)

                    # 1. Scroll to entry
                    try:
                        entry_element = driver.find_element(By.ID, f"epb_entry_{block_id}")
                        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", entry_element)
                        time.sleep(1)
                        print(f"         ✓ Scrolled to entry {block_id}")
                    except Exception as e:
                        print(f"         ⚠️  Could not find entry {block_id}: {e}")

                    # 2. Locate container - use the correct CSS selector
                    type_class = 'el-type-4' if elem_type == 'table' else 'el-type-7'
                    container_selector = f"[id^='container_{block_id}_'].{type_class}"
                    
                    try:
                        containers = driver.find_elements(By.CSS_SELECTOR, container_selector)

                        print(f"         Found {len(containers)} matching containers")

                        if not containers:
                            print(f"         ✗ Could not find container")
                            self.stats['xlsx_failed'] += 1
                            continue
                        
                        # Try each matching container until one yields a successful download
                        download_success = False
                        downloaded_file = None
                        server_filename = None
                        for idx, container in enumerate(containers):
                            try:
                                print(f"         Trying container #{idx} id={container.get_attribute('id')}")

                                # Scroll the internal entries container so the target container is at the top,
                                # then scroll that internal container further down by the configured nudge
                                try:
                                    driver.execute_script(r"""
                                        var el = arguments[0];
                                        var scrollContainer = arguments[1];
                                        var nudge = arguments[2] || 0;
                                        try {
                                            if (scrollContainer) {
                                                var containerRect = scrollContainer.getBoundingClientRect();
                                                var elRect = el.getBoundingClientRect();
                                                var offset = elRect.top - containerRect.top;
                                                // move so the element is positioned at the top of the scroll container
                                                scrollContainer.scrollTop = scrollContainer.scrollTop + offset;
                                                // then scroll further down by nudge to reveal buttons below
                                                scrollContainer.scrollTop = scrollContainer.scrollTop + nudge;
                                            } else {
                                                el.scrollIntoView({block: 'start'});
                                                window.scrollBy(0, nudge);
                                            }
                                        } catch (e) {
                                            try { el.scrollIntoView({block: 'center'}); } catch (ee) {}
                                            try { window.scrollBy(0, nudge); } catch (ee) {}
                                        }
                                    """, container, scroll_container, self.scroll_nudge)
                                except Exception:
                                    try:
                                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", container)
                                        driver.execute_script("window.scrollBy(0, arguments[0]);", self.scroll_nudge)
                                    except:
                                        pass
                                time.sleep(0.5)

                                # visually mark the container (red ring) and bring to top layer so it's clear
                                try:
                                    driver.execute_script("""
                                        var el = arguments[0];
                                        el.style.zIndex = 9999;
                                        el.style.position = 'relative';
                                        el.style.boxShadow = '0 0 0 3px rgba(255,0,0,0.35)';
                                    """, container)
                                except:
                                    pass

                                # Hover to show buttons
                                actions = ActionChains(driver)
                                actions.move_to_element(container).perform()
                                time.sleep(0.5)

                                # Force-show button wrapper (in case CSS didn't trigger)
                                driver.execute_script("""
                                    var container = arguments[0];
                                    var buttonWrapper = container.querySelector('.button_wrapper');
                                    if (buttonWrapper) {
                                        buttonWrapper.style.opacity = '1';
                                        buttonWrapper.style.visibility = 'visible';
                                        buttonWrapper.style.display = 'flex';
                                    }
                                """, container)

                                # Find and click the more-options button
                                try:
                                    button = container.find_element(By.CSS_SELECTOR, 'button.more_options_button')
                                except Exception as e:
                                    print(f"         ⚠️  Container #{idx} has no more_options_button: {e}")
                                    continue

                                actions = ActionChains(driver)
                                actions.move_to_element(button).click().perform()
                                time.sleep(1.5)

                                # Find the Excel download option
                                excel_options = driver.find_elements(By.XPATH, "//span[contains(text(), 'Download as Excel')]")
                                if not excel_options:
                                    excel_options = driver.find_elements(By.XPATH, "//span[contains(text(), 'Excel')]")

                                if not excel_options:
                                    print(f"         ⚠️  Container #{idx}: Excel option not found after clicking menu")
                                    # Try to close menu and continue to next container
                                    try:
                                        actions = ActionChains(driver)
                                        actions.move_by_offset(200, 200).click().perform()
                                    except:
                                        pass
                                    time.sleep(0.5)
                                    continue

                                # Click the visible Excel option
                                visible_option = None
                                for opt in excel_options:
                                    if opt.is_displayed():
                                        visible_option = opt
                                        break

                                try:
                                    if visible_option:
                                        actions = ActionChains(driver)
                                        actions.move_to_element(visible_option).click().perform()
                                    else:
                                        excel_options[0].click()
                                except Exception as e:
                                    print(f"         ⚠️  Clicking Excel option failed for container #{idx}: {e}")
                                    # try next container
                                    continue

                                time.sleep(1.5)

                                # Wait for download to complete (check temp directory)
                                print(f"         Waiting for download to complete (up to 30 seconds)...")
                                timeout = 30
                                start_time = time.time()

                                while time.time() - start_time < timeout:
                                    xlsx_files = list(temp_dir.glob('*.xlsx'))
                                    temp_files = list(temp_dir.glob('*.crdownload')) + list(temp_dir.glob('*.tmp'))
                                    if xlsx_files and not temp_files:
                                        downloaded_file = xlsx_files[0]
                                        server_filename = downloaded_file.name
                                        print(f"         ✓ Download complete: {server_filename}")
                                        time.sleep(1)
                                        download_success = True
                                        break
                                    time.sleep(0.5)

                                if download_success:
                                    # successful for this container
                                    break
                                else:
                                    print(f"         ✗ Container #{idx} timed out waiting for download")
                                    # try next container
                                    continue

                            except Exception as e:
                                print(f"         ⚠️  Error while processing container #{idx}: {e}")
                                continue

                        if not download_success:
                            print(f"         ✗ Could not download XLSX from any matching container")
                            self.stats['xlsx_failed'] += 1
                            continue
                        
                        if downloaded_file and downloaded_file.exists():
                            # Generate target filename: use the server's original filename (already unique)
                            # Format: server_filename = "Labfolder Table.xlsx", "Labfolder Well Plate Template.xlsx", etc.
                            timestamp = int(time.time())

                            # Keep the server filename, add timestamp and element_id to ensure uniqueness
                            server_basename = downloaded_file.stem  # Exclude extension
                            final_filename = f"{elem_type}_{block_id}_{elem_id}_{server_basename}_{timestamp}.xlsx"
                            final_path = entry_dir / final_filename

                            # Immediately move file to target location
                            print(f"         📦 Move: {server_filename} -> {final_filename}")
                            try:
                                shutil.move(str(downloaded_file), str(final_path))
                            except Exception as e:
                                print(f"         ❌ Move file failed: {e}")
                                # 尝试复制然后删除
                                try:
                                    shutil.copy2(str(downloaded_file), str(final_path))
                                    downloaded_file.unlink()
                                except Exception as e2:
                                    print(f"         ❌ Copy file also failed: {e2}")
                                    continue
                            
                            # 验证文件大小
                            file_size = final_path.stat().st_size
                            print(f"         📊 File size: {file_size} bytes")
                            print(f"         ✅ Saved to: {final_path.relative_to(self.out)}")

                            # Read title information from JSON and save filename mapping
                            json_file = entry_dir / f'{elem_type}_{elem_id}.json'
                            elem_title = ""
                            if json_file.exists():
                                try:
                                    with open(json_file, 'r', encoding='utf-8') as f:
                                        data = json.load(f)
                                    elem_title = data.get('title', '')
                                    # Save XLSX filename to JSON
                                    data['xlsx_filename'] = final_filename
                                    data['xlsx_title'] = elem_title  # Save original title for HTML display
                                    data['server_filename'] = server_filename  # Save server original filename for debugging
                                    with open(json_file, 'w', encoding='utf-8') as f:
                                        json.dump(data, f, ensure_ascii=False, indent=2)
                                except Exception as e:
                                    print(f"         ⚠️  Update JSON failed: {e}")

                            self.stats['xlsx_downloaded'] += 1
                            print(f"         ✓ Process completed, continue to next...")
                            time.sleep(2)  # Short wait for 2 seconds
                        else:
                            print(f"         ✗ Download timed out")
                            self.stats['xlsx_failed'] += 1
                    
                    except Exception as e:
                        print(f"         ✗ Error: {e}")
                        self.stats['xlsx_failed'] += 1
                        self.stats['errors'].append(f"XLSX download failed for {elem_type} {elem_id}: {e}")
                
                except Exception as e:
                    print(f"      ✗ Processing failed: {e}")
                    self.stats['xlsx_failed'] += 1
            
            print(f"    ✓ Downloaded: {self.stats['xlsx_downloaded']}")
            print(f"    ✗ Failed: {self.stats['xlsx_failed']}")
        
        finally:
            driver.quit()
            # Clean up temporary directory
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
    
    def generate_entry_html(self, entry):
        """Generate individual entry HTML (including embedded tables and wellplates)"""
        entry_id = entry['id']
        title = entry.get('title', f'Entry {entry_id}')
        
        user_vm = entry.get('userViewModel', {})
        author_name = f"{user_vm.get('firstname', '')} {user_vm.get('lastname', '')}".strip()
        
        created = entry.get('createTS', '')
        modified = entry.get('versionTS', '')
        project = entry.get('projectName', '')
        
        safe_proj = "".join(c for c in project if c.isalnum() or c in (' ','-','_')).strip() or 'Unknown'
        safe_title = "".join(c for c in title if c.isalnum() or c in (' ','-','_')).strip() or f'Entry_{entry_id}'
        entry_dir = self.out / safe_proj / f"{entry_id}_{safe_title}"
        
        content_html = ""
        
        for row in entry.get('elements', []):
            for elem in row:
                elem_type = elem.get('type')
                elem_id = elem.get('elementId')
                
                if elem_type == 1:  # Text
                    text_content = elem.get('content', '')
                    if text_content:
                        content_html += f'<div class="text-element" data-element-id="{elem_id}">{text_content}</div>\n'
                
                elif elem_type == 2:  # Image
                    meta = elem.get('metaData', {})
                    fname = meta.get('fileName', f'image_{elem_id}.png')
                    fname = "".join(c if c.isalnum() or c in (' ','-','_','.') else '_' for c in fname).strip()
                    if not fname or fname == '.':
                        fname = f'image_{elem_id}.png'
                    if '.' not in fname:
                        fname += '.png'
                    
                    layer_hash = meta.get('layerHash')
                    layer_fname = fname.replace('.', '_annotated.') if layer_hash else None
                    
                    img_style = 'max-width: 100%; height: auto;'
                    
                    content_html += f'''<div class="image-element" data-element-id="{elem_id}">
                        <div class="image-container" style="position: relative; display: inline-block; margin: 20px 0;">
                            <img class="image-original" src="{fname}" alt="{escape(fname)}" style="{img_style} display: block; border: 1px solid #ddd; border-radius: 4px;">'''
                    
                    if layer_fname:
                        content_html += f'''
                            <img class="image-layer" src="{layer_fname}" alt="Annotated layer" style="{img_style} position: absolute; top: 0; left: 0; pointer-events: none;">'''
                    
                    content_html += '''
                        </div>
                    </div>\n'''
                
                elif elem_type == 3:  # File
                    meta = elem.get('metaData', {})
                    fname = meta.get('fileName', f'file_{elem_id}')
                    fsize = meta.get('fileSize', 0)
                    
                    size_str = f"{fsize/1024:.1f} KB" if fsize < 1024*1024 else f"{fsize/1024/1024:.1f} MB"
                    
                    content_html += f'''<div class="file-element" data-element-id="{elem_id}">
                        <div class="file-attachment">
                            <span class="file-icon">📎</span>
                            <a href="{fname}" download="{fname}">{escape(fname)}</a>
                            <span class="file-size">({size_str})</span>
                        </div>
                    </div>\n'''
                
                elif elem_type == 4:  # Table
                    table_title = elem.get('title', 'Table')
                    table_file = entry_dir / f'table_{elem_id}.json'

                    # Read JSON to get xlsx filename
                    xlsx_link_html = ''
                    if table_file.exists():
                        with open(table_file, 'r', encoding='utf-8') as f:
                            table_data = json.load(f)
                        
                        xlsx_filename = table_data.get('xlsx_filename')
                        if xlsx_filename and (entry_dir / xlsx_filename).exists():
                            xlsx_link_html = f'''
                            <div style="margin: 10px 0;">
                                <a href="{xlsx_filename}" download="{xlsx_filename}" 
                                   style="display: inline-flex; align-items: center; gap: 8px; padding: 8px 16px; background: #4CAF50; color: white; text-decoration: none; border-radius: 4px; font-weight: 500;">
                                    <span>📊</span>
                                    <span>Download as Excel</span>
                                </a>
                            </div>'''
                    
                    if table_file.exists():
                        with open(table_file, 'r', encoding='utf-8') as f:
                            table_data = json.load(f)
                        
                        table_json_str = json.dumps(table_data.get('content', {}))
                        
                        content_html += f'''<div class="table-element" data-element-id="{elem_id}">
                        <h3 style="color: #2c3e50; margin: 20px 0 10px 0;">📊 {escape(table_title)}</h3>
                        {xlsx_link_html}
                        <div id="table-container-{elem_id}" class="table-container" style="background: white; border: 1px solid #ddd; border-radius: 4px; padding: 10px; margin: 15px 0; min-height: 200px;"></div>
                        <script>
                            (function() {{
                                const tableData = {table_json_str};
                                renderTableAsHTML('table-container-{elem_id}', tableData);
                            }})();
                        </script>
                    </div>\n'''
                
                elif elem_type == 7:  # Well Plate
                    well_plate_title = elem.get('title', 'Well Plate')
                    wellplate_file = entry_dir / f'wellplate_{elem_id}.json'

                    # Read JSON to get xlsx filename
                    xlsx_link_html = ''
                    if wellplate_file.exists():
                        with open(wellplate_file, 'r', encoding='utf-8') as f:
                            wellplate_data = json.load(f)
                        
                        xlsx_filename = wellplate_data.get('xlsx_filename')
                        if xlsx_filename and (entry_dir / xlsx_filename).exists():
                            xlsx_link_html = f'''
                            <div style="margin: 10px 0;">
                                <a href="{xlsx_filename}" download="{xlsx_filename}" 
                                   style="display: inline-flex; align-items: center; gap: 8px; padding: 8px 16px; background: #4CAF50; color: white; text-decoration: none; border-radius: 4px; font-weight: 500;">
                                    <span>🧪</span>
                                    <span>Download as Excel</span>
                                </a>
                            </div>'''
                    
                    if wellplate_file.exists():
                        with open(wellplate_file, 'r', encoding='utf-8') as f:
                            wellplate_data = json.load(f)
                        
                        wellplate_json_str = json.dumps(wellplate_data.get('content', {}))
                        
                        content_html += f'''<div class="wellplate-element" data-element-id="{elem_id}">
                        <h3 style="color: #2c3e50; margin: 20px 0 10px 0;">🧪 {escape(well_plate_title)}</h3>
                        {xlsx_link_html}
                        <div id="wellplate-container-{elem_id}" class="wellplate-container" style="background: white; border: 1px solid #4CAF50; border-radius: 4px; padding: 10px; margin: 15px 0; min-height: 200px;"></div>
                        <script>
                            (function() {{
                                const wellplateData = {wellplate_json_str};
                                renderTableAsHTML('wellplate-container-{elem_id}', wellplateData);
                            }})();
                        </script>
                    </div>\n'''
        
        # Complete html
        html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{escape(title)} - LabFolder</title>
    <style>
    body {{
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        margin: 0;
        padding: 20px;
        background: #f5f5f5;
    }}
    .entry-container {{
        max-width: 1200px;
        margin: 0 auto;
        background: white;
        padding: 40px;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }}
    .entry-header {{
        border-bottom: 2px solid #2fa4be;
        padding-bottom: 20px;
        margin-bottom: 30px;
    }}
    .entry-title {{
        font-size: 32px;
        font-weight: bold;
        color: #002b56;
        margin-bottom: 15px;
    }}
    .entry-meta {{
        display: flex;
        gap: 30px;
        color: #666;
        font-size: 14px;
        flex-wrap: wrap;
    }}
    .labfolder-table, .labfolder-wellplate {{
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
        margin: 10px 0;
    }}
    .labfolder-table th, .labfolder-table td,
    .labfolder-wellplate th, .labfolder-wellplate td {{
        padding: 8px 12px;
        border: 1px solid #dee2e6;
        text-align: left;
    }}
    .labfolder-table th, .labfolder-wellplate th {{
        background: #2fa4be;
        color: white;
        font-weight: 600;
    }}
    </style>
    <script>
    function renderTableAsHTML(containerId, content) {{
        if (!content || !content.sheets) {{
            document.getElementById(containerId).innerHTML = '<p style="color: #e74c3c;">No data available</p>';
            return;
        }}
        
        const container = document.getElementById(containerId);
        let html = '';
        
        const sheets = content.sheets;
        const sheetNames = Object.keys(sheets);
        
        for (const sheetName of sheetNames) {{
            const sheet = sheets[sheetName];
            
            if (!sheet.data || !sheet.data.dataTable) {{
                continue;
            }}
            
            const dataTable = sheet.data.dataTable;
            
            let maxRow = -1;
            let maxCol = -1;
            for (const rowIdx in dataTable) {{
                const row = dataTable[rowIdx];
                maxRow = Math.max(maxRow, parseInt(rowIdx));
                for (const colIdx in row) {{
                    maxCol = Math.max(maxCol, parseInt(colIdx));
                }}
            }}
            
            if (maxRow === -1 || maxCol === -1) {{
                continue;
            }}
            
            if (sheetNames.length > 1) {{
                html += `<h4 style="color: #2fa4be; margin: 15px 0 10px 0;">${{sheetName}}</h4>`;
            }}
            
            html += '<div style="overflow-x: auto;"><table class="labfolder-table">';
            
            for (let r = 0; r <= maxRow; r++) {{
                html += '<tr>';
                const rowData = dataTable[r.toString()] || {{}};
                for (let c = 0; c <= maxCol; c++) {{
                    const cell = rowData[c.toString()] || {{}};
                    const value = cell.value !== undefined ? String(cell.value) : '';
                    
                    let cellStyle = 'border: 1px solid #ddd; padding: 8px;';
                    
                    if (cell.style) {{
                        if (cell.style.backColor) {{
                            cellStyle += `background-color: ${{cell.style.backColor}};`;
                        }}
                        if (cell.style.foreColor) {{
                            cellStyle += `color: ${{cell.style.foreColor}};`;
                        }}
                    }}
                    
                    html += `<td style="${{cellStyle}}">${{value}}</td>`;
                }}
                html += '</tr>';
            }}
            
            html += '</table></div>';
        }}
        
        if (html) {{
            container.innerHTML = html;
        }} else {{
            container.innerHTML = '<p style="color: #999;">Empty table</p>';
        }}
    }}
    </script>
</head>
<body>
    <div class="entry-container">
        <div class="entry-header">
            <div class="entry-title">{escape(title)}</div>
            <div class="entry-meta">
                <div class="meta-item">
                    <span class="meta-label">🏷️ Project:</span>
                    <span>{escape(project)}</span>
                </div>
                <div class="meta-item">
                    <span class="meta-label">👤 Author:</span>
                    <span>{escape(author_name)}</span>
                </div>
                <div class="meta-item">
                    <span class="meta-label">📅 Created:</span>
                    <span>{escape(created)}</span>
                </div>
                <div class="meta-item">
                    <span class="meta-label">🆔 Entry ID:</span>
                    <span>{entry_id}</span>
                </div>
            </div>
        </div>
        
        <div class="entry-content">
            {content_html}
        </div>
    </div>
</body>
</html>'''
        
        return html
    
    def generate_html_reports(self):
        """Generate HTML reports"""
        print("\n[5/8] 🎨 Generating HTML reports...")
        print("    🔨 Creating individual entry pages with embedded tables...")
        
        for entry in tqdm(self.entries, desc="    Creating pages"):
            entry_id = entry['id']
            proj = entry.get('projectName', 'Unknown')
            title = entry.get('title', f'Entry_{entry_id}')
            
            safe_proj = "".join(c for c in proj if c.isalnum() or c in (' ','-','_')).strip()
            safe_title = "".join(c for c in title if c.isalnum() or c in (' ','-','_')).strip() or f'Entry_{entry_id}'
            
            entry_dir = self.out / safe_proj / f"{entry_id}_{safe_title}"
            
            entry_html = self.generate_entry_html(entry)
            html_file = entry_dir / 'index.html'
            
            with open(html_file, 'w', encoding='utf-8') as f:
                f.write(entry_html)
        
        print("    🏠 Creating main index page...")
        self.generate_index_page()
        
        print(f"    ✓ Created {len(self.entries)} entry HTML pages")
        print(f"    ✓ Created index page")
    
    def generate_index_page(self):
        """Generate main index page with modern design"""
        
        entries_by_project = {}
        for entry in self.entries:
            proj = entry.get('projectName', 'Unknown')
            entries_by_project.setdefault(proj, []).append(entry)
        
        projects_html = ""
        
        for proj_name, proj_entries in entries_by_project.items():
            safe_proj = "".join(c for c in proj_name if c.isalnum() or c in (' ','-','_')).strip()
            
            projects_html += f'<div class="project-section">\n'
            projects_html += f'<h2 class="project-title">📁 {escape(proj_name)}</h2>\n'
            projects_html += f'<div class="entries-grid">\n'
            
            for entry in proj_entries:
                entry_id = entry['id']
                title = entry.get('title', f'Entry {entry_id}')
                safe_title = "".join(c for c in title if c.isalnum() or c in (' ','-','_')).strip() or f'Entry_{entry_id}'
                
                # Get metadata
                user_vm = entry.get('userViewModel', {})
                author = f"{user_vm.get('firstname', '')} {user_vm.get('lastname', '')}".strip()
                created = entry.get('createTS', '')
                
                # Count elements
                elements = entry.get('elements', [])
                img_count = sum(1 for row in elements for e in row if e.get('type') == 2)
                file_count = sum(1 for row in elements for e in row if e.get('type') == 3)
                table_count = sum(1 for row in elements for e in row if e.get('type') == 4)
                wellplate_count = sum(1 for row in elements for e in row if e.get('type') == 7)
                
                entry_link = f"{safe_proj}/{entry_id}_{safe_title}/index.html"
                
                projects_html += f'''
                <div class="entry-card">
                    <a href="{entry_link}" class="entry-link">
                        <div class="entry-card-title">{escape(title)}</div>
                        <div class="entry-card-meta">
                            <span>👤 {escape(author)}</span>
                            <span>📅 {escape(created)}</span>
                        </div>
                        <div class="entry-card-badges">
                            {f'<span class="badge">🖼️ {img_count}</span>' if img_count > 0 else ''}
                            {f'<span class="badge">📎 {file_count}</span>' if file_count > 0 else ''}
                            {f'<span class="badge">📊 {table_count}</span>' if table_count > 0 else ''}
                            {f'<span class="badge">🧪 {wellplate_count}</span>' if wellplate_count > 0 else ''}
                        </div>
                    </a>
                </div>
                '''
            
            projects_html += '</div>\n</div>\n'
        
        # Calculate total size
        total_size = sum(f.stat().st_size for f in self.out.rglob('*') if f.is_file())
        size_str = f"{total_size/1024/1024:.1f} MB"
        
        index_html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>LabFolder Complete Backup</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }}
        
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }}
        
        header {{
            background: linear-gradient(135deg, #002b56 0%, #2fa4be 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }}
        
        header h1 {{
            font-size: 48px;
            margin-bottom: 10px;
        }}
        
        header p {{
            font-size: 18px;
            opacity: 0.9;
        }}
        
        .stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 40px;
            background: #f8f9fa;
        }}
        
        .stat-card {{
            background: white;
            padding: 30px;
            border-radius: 15px;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            transition: transform 0.3s, box-shadow 0.3s;
        }}
        
        .stat-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(0,0,0,0.15);
        }}
        
        .stat-card .icon {{
            font-size: 48px;
            margin-bottom: 10px;
        }}
        
        .stat-card .number {{
            font-size: 36px;
            font-weight: bold;
            color: #2fa4be;
            margin-bottom: 5px;
        }}
        
        .stat-card .label {{
            font-size: 14px;
            color: #6c757d;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .content {{
            padding: 40px;
        }}
        
        .project-section {{
            margin-bottom: 50px;
        }}
        
        .project-title {{
            font-size: 32px;
            color: #002b56;
            margin-bottom: 25px;
            padding-bottom: 15px;
            border-bottom: 3px solid #2fa4be;
        }}
        
        .entries-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
            gap: 25px;
        }}
        
        .entry-card {{
            background: white;
            border-radius: 12px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            transition: transform 0.3s, box-shadow 0.3s;
            overflow: hidden;
        }}
        
        .entry-card:hover {{
            transform: translateY(-8px);
            box-shadow: 0 12px 30px rgba(0,0,0,0.2);
        }}
        
        .entry-link {{
            display: block;
            padding: 25px;
            text-decoration: none;
            color: inherit;
        }}
        
        .entry-card-title {{
            font-size: 20px;
            font-weight: 600;
            color: #002b56;
            margin-bottom: 15px;
            line-height: 1.4;
        }}
        
        .entry-card-meta {{
            display: flex;
            flex-direction: column;
            gap: 8px;
            margin-bottom: 15px;
            color: #6c757d;
            font-size: 14px;
        }}
        
        .entry-card-badges {{
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
        }}
        
        .badge {{
            background: #e9ecef;
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 12px;
            color: #495057;
            font-weight: 500;
        }}
        
        footer {{
            background: #f8f9fa;
            padding: 30px;
            text-align: center;
            color: #6c757d;
            border-top: 2px solid #e9ecef;
        }}
        
        footer p {{
            font-size: 16px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>� LabFolder Complete Backup</h1>
            <p>Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </header>
        
        <div class="stats">
            <div class="stat-card">
                <div class="icon">📁</div>
                <div class="number">{self.stats['projects']}</div>
                <div class="label">Projects</div>
            </div>
            <div class="stat-card">
                <div class="icon">📝</div>
                <div class="number">{self.stats['entries']}</div>
                <div class="label">Entries</div>
            </div>
            <div class="stat-card">
                <div class="icon">🖼️</div>
                <div class="number">{self.stats['images']}</div>
                <div class="label">Images</div>
            </div>
            <div class="stat-card">
                <div class="icon">📎</div>
                <div class="number">{self.stats['files']}</div>
                <div class="label">Files</div>
            </div>
            <div class="stat-card">
                <div class="icon">📊</div>
                <div class="number">{self.stats['tables']}</div>
                <div class="label">Tables</div>
            </div>
            <div class="stat-card">
                <div class="icon">🧪</div>
                <div class="number">{self.stats['well_plates']}</div>
                <div class="label">Well Plates</div>
            </div>
            <div class="stat-card">
                <div class="icon">📥</div>
                <div class="number">{self.stats['xlsx_downloaded']}</div>
                <div class="label">XLSX Files</div>
            </div>
        </div>
        
        <div class="content">
            {projects_html}
        </div>
        
        <footer>
            <p>📦 Total backup size: {size_str}</p>
            <p style="margin-top: 10px; font-size: 14px;">✨ Complete offline backup with all data, images, files, and downloadable Excel files</p>
        </footer>
    </div>
</body>
</html>'''
        
        with open(self.out / 'index.html', 'w', encoding='utf-8') as f:
            f.write(index_html)
    
    def calculate_sha256(self, file_path):
        """Calculate SHA256 hash of a file"""
        import hashlib
        sha256_hash = hashlib.sha256()
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    
    def export_eln(self):
        """Export to ELN format following RO-Crate specification"""
        print("\n[6/7] 📋 Exporting to ELN format...")
        
        eln_dir = self.out / 'eln_export'
        eln_dir.mkdir(exist_ok=True)
        
        for entry in tqdm(self.entries, desc="    Creating ELN files"):
            entry_id = entry['id']
            title = entry.get('title', f'Entry_{entry_id}')
            
            # Get author info
            user_vm = entry.get('userViewModel', {})
            author_id = f"./author/{user_vm.get('id', 'unknown')}"
            author_firstname = user_vm.get('firstname', '')
            author_lastname = user_vm.get('lastname', '')
            author_email = user_vm.get('email', '')
            
            # Get dates
            created = entry.get('createTS', '')
            modified = entry.get('versionTS', '')
            
            # Sanitize name
            safe_title = "".join(c for c in title if c.isalnum() or c in (' ','-','_')).strip() or f'Entry_{entry_id}'
            safe_proj = "".join(c for c in entry.get('projectName', 'Unknown') if c.isalnum() or c in (' ','-','_')).strip()
            
            # Create ELN package folder
            eln_pkg_name = f"{entry_id}_{safe_title}"
            eln_pkg_dir = eln_dir / eln_pkg_name
            eln_pkg_dir.mkdir(exist_ok=True)
            
            # Create experiment folder with safe name
            if created:
                # Extract date part and sanitize - handle various date formats
                # Could be: "2023-08-01T15:08:00" or "01.08.2023 15:08" or other formats
                date_part = created.split('T')[0] if 'T' in created else created.split(' ')[0]
                # Replace dots and other separators with hyphens, keep only first 10 chars for date
                date_part = date_part.replace('.', '-').replace('/', '-')[:10]
                exp_folder_name = f"{date_part} - {safe_title}"
            else:
                exp_folder_name = safe_title
            # Remove any remaining special characters but keep hyphens and spaces
            exp_folder_name = "".join(c for c in exp_folder_name if c.isalnum() or c in (' ','-','_')).strip()[:100]
            
            exp_dir = eln_pkg_dir / exp_folder_name
            exp_dir.mkdir(exist_ok=True)
            
            # Build complete HTML content with file mentions in bold
            html_content = ""
            for row in entry.get('elements', []):
                for elem in row:
                    elem_type = elem.get('type')
                    
                    if elem_type == 1:  # Text
                        content = elem.get('content', '')
                        if content:
                            html_content += content + "\n\n"
                    
                    elif elem_type == 2:  # Image
                        meta = elem.get('metaData', {})
                        fname = meta.get('fileName', 'image.png')
                        html_content += f'<p><strong style="color: #1976d2;">📷 Image: {fname}</strong></p>\n'
                    
                    elif elem_type == 3:  # File
                        meta = elem.get('metaData', {})
                        fname = meta.get('fileName', 'file')
                        html_content += f'<p><strong style="color: #388e3c;">📎 File: {fname}</strong></p>\n'
                    
                    elif elem_type == 4:  # Table
                        table_title = elem.get('title', 'Table')
                        html_content += f'<p><strong>📊 Table: {table_title}</strong></p>\n'
                    
                    elif elem_type == 7:  # Well Plate
                        wp_title = elem.get('title', 'Well Plate')
                        html_content += f'<p><strong>🧪 Well Plate: {wp_title}</strong></p>\n'
            
            # Initialize graph
            graph = []
            
            # 1. Metadata descriptor
            graph.append({
                "@id": "ro-crate-metadata.json",
                "@type": "CreativeWork",
                "about": {"@id": "./"},
                "conformsTo": {"@id": "https://w3id.org/ro/crate/1.1"},
                "dateCreated": datetime.now().isoformat()
            })
            
            # 2. Root directory
            graph.append({
                "@id": "./",
                "@type": "Dataset",
                "hasPart": [{"@id": f"./{exp_folder_name}/"}]
            })
            
            # 3. Author
            graph.append({
                "@id": author_id,
                "@type": "Person",
                "givenName": author_firstname,
                "familyName": author_lastname,
                "email": author_email
            })
            
            # 4. Experiment Dataset
            exp_parts = []
            
            # Copy files
            source_entry_dir = self.out / safe_proj / f"{entry_id}_{safe_title}"
            
            if source_entry_dir.exists():
                for src_file in source_entry_dir.iterdir():
                    if not src_file.is_file():
                        continue
                    if src_file.name in ['entry_data.json', 'index.html']:
                        continue
                    # Skip text files
                    if src_file.name.startswith('text_'):
                        continue
                    
                    try:
                        # Shorten filename to avoid Windows path length issues
                        # Use first 20 chars + timestamp for uniqueness
                        original_name = src_file.name
                        
                        # Split name and extension
                        name_parts = original_name.rsplit('.', 1)
                        if len(name_parts) == 2:
                            base_name, extension = name_parts
                        else:
                            base_name = original_name
                            extension = ''
                        
                        # Create short, unique filename: first 20 chars + timestamp
                        import time
                        import platform
                        timestamp = str(int(time.time() * 1000))[-10:]  # Last 10 digits of millisecond timestamp
                        short_base = base_name[:20]
                        dest_name = f"{short_base}_{timestamp}.{extension}" if extension else f"{short_base}_{timestamp}"
                        
                        # On Windows, notify user if filename is shortened
                        if platform.system() == 'Windows' and len(original_name) > 50:
                            print(f"      📝 Renamed (Windows path limit): {original_name[:50]}... → {dest_name}")
                        
                        dest_file = exp_dir / dest_name
                        
                        # Use Windows long path prefix for paths > 200 chars to handle long paths
                        src_path_str = str(src_file.resolve())
                        dest_path_str = str(dest_file.resolve())
                        
                        # Add \\?\ prefix for Windows long path support if needed
                        if len(src_path_str) > 200:
                            if not src_path_str.startswith('\\\\?\\'):
                                src_path_str = '\\\\?\\' + src_path_str
                        if len(dest_path_str) > 200:
                            if not dest_path_str.startswith('\\\\?\\'):
                                dest_path_str = '\\\\?\\' + dest_path_str
                        
                        shutil.copy2(src_path_str, dest_path_str)
                        
                        # Add file to graph
                        file_size = dest_file.stat().st_size
                        file_hash = self.calculate_sha256(dest_file)
                        
                        ext = dest_file.suffix.lower()
                        format_map = {
                            '.png': 'image/png',
                            '.jpg': 'image/jpeg',
                            '.jpeg': 'image/jpeg',
                            '.gif': 'image/gif',
                            '.webp': 'image/webp',
                            '.pdf': 'application/pdf',
                            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                            '.json': 'application/json'
                        }
                        encoding_format = format_map.get(ext, 'application/octet-stream')
                        
                        file_id = f"./{exp_folder_name}/{dest_file.name}"
                        exp_parts.append({"@id": file_id})
                        
                        graph.append({
                            "@id": file_id,
                            "@type": "File",
                            "name": dest_file.name,
                            "encodingFormat": encoding_format,
                            "contentSize": str(file_size),
                            "sha256": file_hash
                        })
                    
                    except Exception as e:
                        print(f"      ⚠️  Warning: Failed to copy {src_file.name}: {e}")
                        continue
            
            # Experiment dataset
            exp_dataset = {
                "@id": f"./{exp_folder_name}/",
                "@type": "Dataset",
                "author": {"@id": author_id},
                "name": title,
                "identifier": str(entry_id),
                "dateCreated": created,
                "dateModified": modified,
                "text": html_content,
            }
            
            if exp_parts:
                exp_dataset["hasPart"] = exp_parts
            
            if entry.get('projectName'):
                exp_dataset["keywords"] = [entry['projectName']]
            
            graph.append(exp_dataset)
            
            # Create ro-crate-metadata.json
            ro_crate = {
                "@context": "https://w3id.org/ro/crate/1.1/context",
                "@graph": graph
            }
            
            metadata_file = eln_pkg_dir / "ro-crate-metadata.json"
            with open(metadata_file, 'w', encoding='utf-8') as f:
                json.dump(ro_crate, f, indent=2, ensure_ascii=False)
            
            # Create ZIP with .eln extension
            eln_file = eln_dir / f"{eln_pkg_name}.eln"
            with zipfile.ZipFile(eln_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, dirs, files in os.walk(eln_pkg_dir):
                    for file in files:
                        file_path = Path(root) / file
                        arcname = file_path.relative_to(eln_pkg_dir.parent)
                        try:
                            zf.write(str(file_path), str(arcname))
                        except Exception as e:
                            print(f"      ⚠️  Warning: Failed to add {file[:30]}... to ZIP: {e}")
            
            # Remove temp folder
            try:
                shutil.rmtree(eln_pkg_dir)
            except Exception as e:
                print(f"      ⚠️  Warning: Could not remove temp directory: {e}")
        
        print(f"    ✓ Created {len(self.entries)} ELN packages (RO-Crate compliant)")
        print(f"    ✓ Files listed in text as bold labels, actual files as attachments")
    
    def print_summary(self):
        """Print final summary"""
        print("\n[8/8] ✨ Backup Complete!")
        print("\n" + "="*60)
        print("╔" + "═"*58 + "╗")
        print("║" + " "*15 + "🎉 BACKUP SUCCESSFUL! 🎉" + " "*19 + "║")
        print("╚" + "═"*58 + "╝")
        print("\n📊 Final Statistics:")
        print("━"*60)
        print(f"  📁 Projects:          {self.stats['projects']:>6}")
        print(f"  📄 Entries:           {self.stats['entries']:>6}")
        print("  " + "─"*56)
        print(f"  📝 Text blocks:       {self.stats['text_elements']:>6}")
        print(f"  🖼️  Images:            {self.stats['images']:>6}")
        print(f"  📎 Files:             {self.stats['files']:>6}")
        print(f"  📊 Tables:            {self.stats['tables']:>6}")
        print(f"  🧪 Well Plates:       {self.stats['well_plates']:>6}")
        print(f"  📈 Data elements:     {self.stats['data_elements']:>6}")
        print(f"  ✏️  Sketches:          {self.stats['sketches']:>6}")
        print("  " + "─"*56)
        print(f"  📥 XLSX downloaded:   {self.stats['xlsx_downloaded']:>6}")
        if self.stats['xlsx_failed'] > 0:
            print(f"  ⚠️  XLSX failed:       {self.stats['xlsx_failed']:>6}")
        print("━"*60)
        
        # Calculate total size
        total_size = sum(f.stat().st_size for f in self.out.rglob('*') if f.is_file())
        size_mb = total_size / (1024 * 1024)
        size_gb = total_size / (1024 * 1024 * 1024)
        size_str = f"{size_gb:.2f} GB" if size_gb >= 1 else f"{size_mb:.1f} MB"
        
        print(f"\n� Backup Location:")
        print(f"  💾 Total size:        {size_str}")
        print(f"  📍 Directory:         {self.out.absolute()}")
        print(f"  🌐 Index page:        {(self.out / 'index.html').absolute()}")
        print(f"  📋 ELN exports:       {(self.out / 'eln_export').absolute()}")
        
        if self.stats['errors']:
            print(f"\n⚠️  Warnings: {len(self.stats['errors'])} non-critical errors")
            print("  (Backup should still be usable)")
        
        print("\n" + "="*60)
        print("✅ All steps completed successfully!")
        print("\n💡 Next Steps:")
        print("  1. Open index.html in browser")
        print("  2. Check eln_export folder for ELN files")
        print("  3. Verify XLSX files in entry folders")
        print("  4. Backup to external storage")
        print("\n🙏 Thank you for using LabFolder Backup Tool. Use at your own risk.")
        print("="*60)
    
    def run(self):
        """Run complete backup with XLSX download using ActionChains"""
        print("\n" + "="*60)
        print("╔" + "═"*58 + "╗")
        print("║" + " "*10 + "🚀 LabFolder Complete Backup Tool V4" + " "*12 + "║")
        print("║" + " "*58 + "║")
        print("║  📋 Full Backup: Text, Files, Images, Tables" + " "*13 + "║")
        print("║  📊 XLSX Export: Browser automation" + " "*22 + "║")
        print("║  📦 ELN Export: RO-Crate 1.1 compliant" + " "*19 + "║")
        print("║  🎨 HTML Reports: Beautiful web interface" + " "*16 + "║")
        print("╚" + "═"*58 + "╝")
        print("\n" + "─"*60)
        print("📝 Author: Jie Hua")
        print("🏛️ Institution: Imhof Group, BMC, LMU Munich")
        print("📅 Version: 2025-10-26")
        print("─"*60)
        print("\n⚖️  DISCLAIMER:")
        print("   This tool is for LMU Medical Faculty's LabFolder.")
        print("   Other institutions need API modifications.")
        print("   Use at own risk！！！")
        print("\n💡 TIP: Automated process. Grab a coffee! ☕")
        print("   DO NOT interact with browser during automation.")
        print("="*60 + "\n")
        
        try:
            if not self.login():
                return False
            
            self.fetch_data()
            
            print("\n[2.5/8] 🔐 Capturing Authorization Token...")
            self.capture_auth_token()
            
            self.process_entries()
            self.download_xlsx_files()
            self.generate_html_reports()
            self.export_eln()
            
            print("\n[7/8] ✨ Finishing up...")
            self.print_summary()
            
            return True
            
        except KeyboardInterrupt:
            print("\n\n⚠️  User interrupted the process (Ctrl+C), exiting...")
            return False
        except Exception as e:
            print(f"\n\n❌ Error: {e}")
            import traceback
            traceback.print_exc()
            return False

if __name__ == '__main__':
    backup = LabFolderBackup()
    success = backup.run()
    
    if success:
        print("\n✅ Done! Open index.html to view your backup.")
        print("💡 XLSX files are available for download in each entry page.")
    
    sys.exit(0 if success else 1)
