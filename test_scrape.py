import requests
import re

def test():
    url = "https://drive.google.com/drive/folders/1QM4SwF7Tqs3e5-I4kjRMG5ZCUbKGBIZi"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"
    }
    
    print(f"Fetching public Google Drive folder: {url}")
    try:
        r = requests.get(url, headers=headers, timeout=15)
        print(f"Response code: {r.status_code}")
        html = r.text
        print(f"HTML length: {len(html)}")
        
        # Look for any matches of standard Korean file name indicators like '.mp4' or '.mp3' or Korean letters
        # Public Google Drive folder HTML contains file info inside JSON objects in script tags.
        # Let's search for filenames in the HTML using regular expressions.
        # Usually, filenames appear in patterns like:
        # ["1.월.찬양_제목.mp4", ...] or in a JSON structure.
        # Let's try to extract any string containing '월.' or '화.' or '수.' or '목.' or '금.' or '.mp4' or '.mp3'
        
        print("\nSearching for file name patterns...")
        # Search for any strings starting with a number and weekday, like 1.월, 2.화, 3.수, 4.목, 5.금
        # Note: inside JSON strings, quotes are escaped or unicode is used
        patterns = [
            r'1\.월\.[^"]+',
            r'2\.화\.[^"]+',
            r'3\.수\.[^"]+',
            r'4\.목\.[^"]+',
            r'5\.금\.[^"]+'
        ]
        
        found_any = False
        for p in patterns:
            matches = re.findall(p, html)
            if matches:
                print(f"Matched pattern '{p}': {len(matches)} matches")
                for m in matches[:5]:
                    # Clean up unicode escapes if any (e.g. \u003d)
                    try:
                        clean_m = bytes(m, "utf-8").decode("unicode_escape")
                    except Exception:
                        clean_m = m
                    print(f"  - {clean_m}")
                found_any = True
                
        if not found_any:
            print("No simple pattern matches found. Let's do a broader search.")
            # Search for anything matching unicode sequences or containing "mp4" or "mp3" or "wav"
            media_matches = re.findall(r'[^"\\,\s>]*\.(?:mp4|mp3|m4a|wav|avi|mkv)', html)
            print(f"Found media extension matches: {len(media_matches)}")
            for m in list(set(media_matches))[:10]:
                print(f"  - {m}")
                
            # Let's print out the script tags that contain JSON or initial data to see if we can find them there
            # Usually, there's a script tag starting with 'window._initialData' or '_INITIAL_STATE_' or 'ytInitialData'
            initial_state_match = re.search(r'INITIAL_STATE\s*=\s*(\{.*?\});', html)
            if initial_state_match:
                print("Found _INITIAL_STATE_ JSON!")
            else:
                print("Could not find _INITIAL_STATE_.")
                
            # Let's try to inspect raw strings inside the HTML matching Korean characters or "찬양"
            chanyang_matches = re.findall(r'[^"\\,\s]*찬양[^"\\,\s]*', html)
            print(f"Found matches containing '찬양': {len(chanyang_matches)}")
            for m in list(set(chanyang_matches))[:15]:
                print(f"  - {m}")
                
    except Exception as e:
        print(f"Error occurred: {str(e)}")

if __name__ == '__main__':
    test()
