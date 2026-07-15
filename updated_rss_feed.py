import feedparser
from datetime import datetime, timedelta
import json
import sys
import re
import requests
import ssl
import os

# The Amazon SP-API RSS Feed URL - Using HTTP (works in browser)
RSS_URL = "http://developer-docs.amazon.com/sp-api/changelog.rss"

# Browser headers to avoid being blocked
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'application/rss+xml, application/xml, text/xml, */*',
    'Accept-Language': 'en-US,en;q=0.9',
}

# State file to store previously seen GUIDs
STATE_FILE = "rss_state.json"
# Output text file for email
OUTPUT_FILE = "amazon_rss_update.txt"

def fetch_rss_feed():
    """
    Fetch and parse the RSS feed using requests
    """
    print("📡 Fetching RSS feed from Amazon...")
    
    try:
        # Use requests to fetch the feed
        response = requests.get(RSS_URL, headers=HEADERS, timeout=30)
        
        print(f"    Status: {response.status_code}")
        
        if response.status_code != 200:
            print(f"❌ HTTP Error: {response.status_code}")
            return None
        
        print(f"    Downloaded: {len(response.content):,} bytes")
        
        # Parse with feedparser
        print("    Parsing RSS feed...")
        feed = feedparser.parse(response.content)
        
        if feed.bozo:
            print(f"    Warning: {feed.bozo_exception}")
        
        if feed.entries:
            print(f"✅ Success! Found {len(feed.entries)} items")
            return feed
        else:
            print("    No entries found")
            return None
            
    except requests.exceptions.RequestException as e:
        print(f"❌ Request Error: {e}")
        return None
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

def fetch_rss_https():
    """
    Fallback: Try HTTPS with SSL verification disabled
    """
    print("\n🔄 Trying HTTPS fallback...")
    https_url = "https://developer-docs.amazon.com/sp-api/changelog.rss"
    
    try:
        # Disable SSL verification (only as fallback)
        response = requests.get(https_url, headers=HEADERS, verify=False, timeout=30)
        
        if response.status_code == 200:
            feed = feedparser.parse(response.content)
            if feed.entries:
                print(f"✅ HTTPS success! Found {len(feed.entries)} items")
                return feed
        
        return None
        
    except Exception as e:
        print(f"❌ HTTPS fallback failed: {e}")
        return None

def load_state():
    """
    Load previously seen GUIDs from state file
    """
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                print(f"📂 Loaded state: {len(data.get('guids', []))} previously seen GUIDs")
                return data.get('guids', [])
        except Exception as e:
            print(f"⚠️ Error loading state file: {e}")
            return []
    else:
        print("ℹ️ No state file found. This is the first run.")
        return []

def save_state(guids, total_items, last_build_date):
    """
    Save current GUIDs to state file
    """
    try:
        data = {
            "last_check": datetime.now().isoformat(),
            "last_build_date": last_build_date,
            "total_items": total_items,
            "guids": guids
        }
        with open(STATE_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"💾 State saved: {len(guids)} GUIDs")
        return True
    except Exception as e:
        print(f"❌ Error saving state: {e}")
        return False
     
def extract_guid(entry):
    """
    Extract GUID from RSS entry (handles different formats)
    """
    guid = entry.get('id') or entry.get('guid') or entry.get('link')
    
    # Handle if guid is an object with a 'value' attribute
    if hasattr(guid, 'value'):
        guid = guid.value
    
    # If it's a link and no proper GUID, use the link as fallback
    if not guid:
        guid = entry.get('link', '')
    
    return str(guid)

def find_new_items(feed, old_guids):
    """
    Find items that are in the feed but NOT in old_guids
    """
    current_guids = []
    new_items = []
    
    for entry in feed.entries:
        guid = extract_guid(entry)
        current_guids.append(guid)
        
        if guid not in old_guids:
            new_items.append({
                'guid': guid,
                'title': entry.get('title', 'No Title'),
                'link': entry.get('link', ''),
                'pub_date': entry.get('published', 'Unknown date'),
                'description': entry.get('description', 'No description'),
                'author': entry.get('author', ''),
                'raw_entry': entry
            })
    
    return current_guids, new_items

def build_text_file(feed, new_items, current_guids):
    """
    Build the text file with:
    - Header
    - PART 1: New items
    - PART 2: All items
    """
    lines = []
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # ============================================================
    # HEADER
    # ============================================================
    lines.append("=" * 80)
    lines.append("AMAZON SP-API CHANGELOG UPDATE")
    lines.append("=" * 80)
    lines.append(f"Generated: {now}")
    lines.append(f"Feed Last Build Date: {feed.feed.get('lastBuildDate', 'Unknown')}")
    lines.append(f"Total Items in Feed: {len(feed.entries)}")
    lines.append("")
    lines.append(f"Source: {RSS_URL}")
    lines.append("=" * 80)
    lines.append("")
    
    # ============================================================
    # PART 1: NEW UPDATES (Since Last Check)
    # ============================================================
    lines.append("=" * 80)
    lines.append("PART 1: NEW UPDATES (Since Last Check)")
    lines.append("=" * 80)
    
    if new_items:
        lines.append(f"Total: {len(new_items)} new item(s)")
        lines.append("")
        
        for idx, item in enumerate(new_items, 1):
            lines.append(f"{'─' * 40}")
            lines.append(f"ITEM #{idx}")
            lines.append(f"Title: {item['title']}")
            lines.append(f"Date: {item['pub_date']}")
            lines.append(f"Link: {item['link']}")
            lines.append(f"GUID: {item['guid']}")
            if item['author']:
                lines.append(f"Author: {item['author']}")
            # Add description (first 300 chars)
            desc = item['description']
            if len(desc) > 300:
                desc = desc[:300] + "..."
            lines.append(f"Description: {desc}")
            lines.append("")
    else:
        lines.append("✅ No new items found. All updates have been previously sent.")
        lines.append("")
    
    # ============================================================
    # PART 2: COMPLETE RSS FEED (All Current Items)
    # ============================================================
    lines.append("=" * 80)
    lines.append("PART 2: COMPLETE RSS FEED (All Current Items)")
    lines.append("=" * 80)
    lines.append(f"Total: {len(feed.entries)} item(s)")
    lines.append("")
    
    for idx, entry in enumerate(feed.entries, 1):
        guid = extract_guid(entry)
        # Check if this is a new item (mark with *)
        is_new = "⭐ " if guid in [item['guid'] for item in new_items] else "   "
        
        lines.append(f"{is_new}{idx}. {entry.get('title', 'No Title')}")
        lines.append(f"   Date: {entry.get('published', 'Unknown date')}")
        lines.append(f"   GUID: {guid}")
        lines.append(f"   Link: {entry.get('link', '')}")
        # Add description summary (first 150 chars)
        desc = entry.get('description', '')
        #if len(desc) > 150:
        #    desc = desc[:150] + "..."
        if desc:
            lines.append(f"   Summary: {desc}")
        lines.append("")
    
    # ============================================================
    # FOOTER
    # ============================================================
    lines.append("=" * 80)
    lines.append("END OF REPORT")
    lines.append("=" * 80)
    lines.append("")
    lines.append(f"Report generated by Amazon RSS Monitor")
    lines.append(f"© {datetime.now().year}")
    
    return "\n".join(lines)

def main():
    """
    Main function
    """
    print("\n" + "=" * 80)
    print("📊 AMAZON RSS FEED MONITOR WITH TRACKING")
    print("=" * 80)
    
    # Check virtual environment
    in_venv = hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix)
    if in_venv:
        print("✅ Running in virtual environment")
    else:
        print("⚠️ Not in virtual environment")
    
    # Try HTTP first (works in browser)
    feed = fetch_rss_feed()
    
    # If HTTP fails, try HTTPS
    if feed is None or not feed.entries:
        feed = fetch_rss_https()
    
    # If still no entries, try direct feedparser
    if feed is None or not feed.entries:
        print("\n🔄 Trying direct feedparser...")
        try:
            feed = feedparser.parse(RSS_URL)
            if feed.entries:
                print(f"✅ Direct feedparser found {len(feed.entries)} items")
        except Exception as e:
            print(f"❌ Direct feedparser failed: {e}")
    
    # Check if we got anything
    if feed is None or not feed.entries:
        print("\n❌ Failed to fetch RSS feed.")
        print("\nTroubleshooting steps:")
        print("1. Check your internet connection")
        print("2. Try accessing in browser: http://developer-docs.amazon.com/sp-api/changelog.rss")
        print("3. If browser works, the issue is with Python SSL certificates")
        print("   - Try: pip install --upgrade certifi")
        return
    
    # ============================================================
    # MAIN LOGIC: Compare and Build Report
    # ============================================================
    print("\n" + "=" * 80)
    print("📋 ANALYZING FEED")
    print("=" * 80)
    
    # Load previous state
    old_guids = load_state()
    
    # Find new items
    current_guids, new_items = find_new_items(feed, old_guids)
    
    # Display what was found
    print(f"\n📊 Current feed: {len(current_guids)} items")
    print(f"📊 Previously seen: {len(old_guids)} items")
    print(f"🆕 New items found: {len(new_items)}")
    
    if new_items:
        print("\n🆕 New items:")
        for item in new_items:
            print(f"   • {item['title'][:80]}...")
            print(f"     Date: {item['pub_date']}")
            print(f"     GUID: {item['guid']}")
    
    # Build the text file
    print("\n📝 Building text file...")
    text_content = build_text_file(feed, new_items, current_guids)
    
    # Save the text file
    try:
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            f.write(text_content)
        print(f"✅ Text file saved: {OUTPUT_FILE}")
        print(f"   File size: {len(text_content):,} characters")
        print(f"   Lines: {len(text_content.splitlines()):,}")
    except Exception as e:
        print(f"❌ Error saving text file: {e}")
    
    # Update and save state
    print("\n💾 Updating state...")
    last_build_date = feed.feed.get('lastBuildDate', datetime.now().isoformat())
    if save_state(current_guids, len(current_guids), last_build_date):
        print("✅ State updated successfully")
    
    # Summary
    print("\n" + "=" * 80)
    print("✅ COMPLETED")
    print("=" * 80)
    print(f"\n📁 Files created/updated:")
    print(f"   - {OUTPUT_FILE} (Email attachment)")
    print(f"   - {STATE_FILE} (Tracking state)")
    
    if new_items:
        print(f"\n📧 Ready to email: {len(new_items)} new items found")
        print(f"   Attach: {OUTPUT_FILE}")
    else:
        print(f"\nℹ️  No new items. You can still send the complete feed.")
        print(f"   Attach: {OUTPUT_FILE}")
    
    print("\n" + "=" * 80)

if __name__ == "__main__":
    # Suppress SSL warnings when using verify=False
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    main()