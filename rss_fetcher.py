import feedparser
from datetime import datetime, timedelta
import json
import sys
import re
import requests
import ssl

# The Amazon SP-API RSS Feed URL - Using HTTP (works in browser)
RSS_URL = "http://developer-docs.amazon.com/sp-api/changelog.rss"

# Browser headers to avoid being blocked
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'application/rss+xml, application/xml, text/xml, */*',
    'Accept-Language': 'en-US,en;q=0.9',
}

def fetch_rss_feed():
    """
    Fetch and parse the RSS feed using requests
    """
    print("🔄 Fetching RSS feed from Amazon...")
    
    try:
        # Use requests to fetch the feed
        response = requests.get(RSS_URL, headers=HEADERS, timeout=30)
        
        print(f"   📡 Status: {response.status_code}")
        
        if response.status_code != 200:
            print(f"❌ HTTP Error: {response.status_code}")
            return None
        
        print(f"   📦 Downloaded: {len(response.content):,} bytes")
        
        # Parse with feedparser
        print("   🔄 Parsing RSS feed...")
        feed = feedparser.parse(response.content)
        
        if feed.bozo:
            print(f"   ⚠️ Warning: {feed.bozo_exception}")
        
        if feed.entries:
            print(f"   ✅ Success! Found {len(feed.entries)} items")
            return feed
        else:
            print("   ❌ No entries found")
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
                print(f"   ✅ HTTPS success! Found {len(feed.entries)} items")
                return feed
        
        return None
        
    except Exception as e:
        print(f"   ❌ HTTPS fallback failed: {e}")
        return None

def display_feed_info(feed):
    """
    Display basic information about the feed
    """
    print("\n" + "="*80)
    print("📰 AMAZON SP-API CHANGELOG FEED")
    print("="*80)
    
    print(f"\n📌 Feed Title: {feed.feed.get('title', 'N/A')}")
    print(f"📌 Last Updated: {feed.feed.get('lastBuildDate', 'N/A')}")
    print(f"📌 Total Items: {len(feed.entries)}")
    print("\n" + "-"*80)

def categorize_items(feed):
    """
    Categorize items by type
    """
    categories = {
        "🆕 NEW APIs/Features": [],
        "🔴 DEPRECATIONS/Removals": [],
        "🟡 UPDATES/Changes": [],
        "📋 POLICY/Agreement": [],
        "ℹ️ OTHER": []
    }
    
    keywords = {
        "🆕 NEW APIs/Features": ['new', 'introducing', 'launch', 'announce'],
        "🔴 DEPRECATIONS/Removals": ['deprecat', 'remov', 'end', 'discontinu', 'sunset', 'final'],
        "🟡 UPDATES/Changes": ['update', 'change', 'effective', 'migrat'],
        "📋 POLICY/Agreement": ['policy', 'agreement', 'data protection', 'acceptable use']
    }
    
    for entry in feed.entries:
        text = (entry.title + " " + entry.description).lower()
        categorized = False
        
        for category, words in keywords.items():
            if any(word in text for word in words):
                categories[category].append(entry)
                categorized = True
                break
        
        if not categorized:
            categories["ℹ️ OTHER"].append(entry)
    
    return categories

def display_summary_by_category(categories):
    """
    Display categorized summary
    """
    print("\n" + "="*80)
    print("📊 CATEGORIZED SUMMARY")
    print("="*80)
    
    for category, items in categories.items():
        if items:
            print(f"\n{category}: {len(items)} items")
            for item in items[:3]:
                pub_date = item.get('published', 'Unknown date')
                print(f"   • {item.title[:80]}...")
                print(f"     📅 {pub_date}")
            if len(items) > 3:
                print(f"   ... and {len(items)-3} more")

def get_latest_items(feed, limit=10):
    """
    Get the most recent items
    """
    print("\n" + "="*80)
    print(f"🆕 LATEST {limit} ITEMS")
    print("="*80)
    
    for idx, entry in enumerate(feed.entries[:limit], 1):
        pub_date = entry.get('published', 'Unknown date')
        print(f"\n{idx}. {entry.title}")
        print(f"   📅 {pub_date}")
        print(f"   📄 {entry.description[:150]}...")
        print(f"   🔗 {entry.link}")

def search_items(feed, search_term):
    """
    Search for specific terms in the feed
    """
    found = []
    for entry in feed.entries:
        if search_term.lower() in (entry.title + entry.description).lower():
            found.append(entry)
    
    if found:
        print(f"\nFound {len(found)} items:")
        for idx, entry in enumerate(found[:5], 1):
            print(f"   {idx}. {entry.title[:100]}...")
            print(f"      📅 {entry.get('published', 'Unknown')}")
    else:
        print(f"\n❌ No items found")

def get_critical_updates(feed):
    """
    Get critical updates (deprecations, removals, endings)
    """
    critical = []
    keywords = ['deprecat', 'remov', 'end', 'discontinu', 'sunset', 'final', 'will be removed', 'no longer']
    
    for entry in feed.entries:
        text = (entry.title + " " + entry.description).lower()
        if any(kw in text for kw in keywords):
            critical.append(entry)
    
    return critical

def get_upcoming_deadlines(feed):
    """
    Find items with upcoming deadlines
    """
    deadlines = []
    date_pattern = r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b'
    
    for entry in feed.entries:
        dates = re.findall(date_pattern, entry.description)
        if dates:
            deadlines.append({
                'title': entry.title,
                'dates': dates,
                'link': entry.link,
                'published': entry.get('published', 'Unknown')
            })
    
    return deadlines

def save_to_json(feed, filename="rss_data.json"):
    """
    Save feed data to JSON file
    """
    data = {
        "feed_info": {
            "title": feed.feed.get('title', 'N/A'),
            "last_build_date": feed.feed.get('lastBuildDate', 'N/A'),
            "total_items": len(feed.entries),
            "fetched_at": datetime.now().isoformat()
        },
        "items": []
    }
    
    for entry in feed.entries:
        data["items"].append({
            "title": entry.title,
            "description": entry.description,
            "link": entry.link,
            "published_date": entry.get('published', ''),
            "author": entry.get('dc:creator', '')
        })
    
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    print(f"\n✅ Data saved to {filename}")

def generate_email_report(feed):
    """
    Generate a formatted email report
    """
    critical = get_critical_updates(feed)
    
    report = []
    report.append("="*80)
    report.append("AMAZON SP-API UPDATE SUMMARY")
    report.append("="*80)
    report.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report.append(f"Total Updates: {len(feed.entries)}")
    report.append("")
    
    # Critical updates
    report.append("🔴 CRITICAL UPDATES (Deprecations/Removals)")
    report.append("-"*40)
    if critical:
        for item in critical[:10]:
            report.append(f"• {item.title}")
            report.append(f"  {item.get('published', 'Unknown date')}")
            report.append(f"  {item.link}")
            report.append("")
    else:
        report.append("No critical updates found.")
    
    # Latest 5 updates
    report.append("")
    report.append("🆕 LATEST 5 UPDATES")
    report.append("-"*40)
    for item in feed.entries[:5]:
        report.append(f"• {item.title}")
        report.append(f"  {item.get('published', 'Unknown date')}")
        report.append("")
    
    return "\n".join(report)

def main():
    """
    Main function
    """
    print("\n" + "="*80)
    print("🚀 AMAZON RSS FEED MONITOR")
    print("="*80)
    
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
                print(f"   ✅ Direct feedparser found {len(feed.entries)} items")
        except Exception as e:
            print(f"   ❌ Direct feedparser failed: {e}")
    
    # Check if we got anything
    if feed is None or not feed.entries:
        print("\n❌ Failed to fetch RSS feed.")
        print("\nTroubleshooting steps:")
        print("1. Check your internet connection")
        print("2. Try accessing in browser: http://developer-docs.amazon.com/sp-api/changelog.rss")
        print("3. If browser works, the issue is with Python SSL certificates")
        print("   - Try: pip install --upgrade certifi")
        return
    
    # Display all the information
    display_feed_info(feed)
    
    # Categorized summary
    categories = categorize_items(feed)
    display_summary_by_category(categories)
    
    # Latest items
    get_latest_items(feed, limit=10)
    
    # Critical updates
    print("\n" + "="*80)
    print("🔴 CRITICAL UPDATES (Deprecations/Removals)")
    print("="*80)
    critical = get_critical_updates(feed)
    if critical:
        print(f"\nFound {len(critical)} critical updates:")
        for idx, item in enumerate(critical[:10], 1):
            print(f"\n{idx}. {item.title}")
            print(f"   📅 {item.get('published', 'Unknown')}")
            print(f"   🔗 {item.link}")
        if len(critical) > 10:
            print(f"\n... and {len(critical)-10} more")
    else:
        print("\nNo critical updates found.")
    
    # Upcoming deadlines
    print("\n" + "="*80)
    print("📅 UPDATES WITH DEADLINES")
    print("="*80)
    deadlines = get_upcoming_deadlines(feed)
    if deadlines:
        print(f"\nFound {len(deadlines)} items with upcoming dates:")
        for idx, item in enumerate(deadlines[:10], 1):
            print(f"\n{idx}. {item['title']}")
            print(f"   📅 Dates: {', '.join(item['dates'])}")
            print(f"   🔗 {item['link']}")
    else:
        print("\nNo deadlines found.")
    
    # Search for specific terms
    print("\n" + "="*80)
    print("🔎 SEARCH RESULTS")
    print("="*80)
    
    search_terms = ["settlement report", "orders api", "fulfillment inbound", "listing attribute"]
    for term in search_terms:
        print(f"\n📌 '{term}':")
        search_items(feed, term)
    
    # Save to JSON
    save_to_json(feed)
    
    # Generate email report
    report = generate_email_report(feed)
    with open("email_report.txt", 'w', encoding='utf-8') as f:
        f.write(report)
    print("✅ Email report saved to email_report.txt")
    
    print("\n" + "="*80)
    print("✅ Done! Files created:")
    print("   - rss_data.json (Full data)")
    print("   - email_report.txt (Email-ready summary)")
    print("="*80)

if __name__ == "__main__":
    # Suppress SSL warnings when using verify=False
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    main()