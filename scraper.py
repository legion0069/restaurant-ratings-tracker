import os
import requests
import json
import re
from bs4 import BeautifulSoup
from config import (
    RESTAURANTS_CONFIG,
    GOOGLE_PLACES_API_KEY,
    SERPAPI_KEY
)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

def get_swiggy_ratings(search_name, restaurant_id=None, lat=17.4400802, lng=78.3809632):
    """
    Fetch rating and total review count from Swiggy's DAPI.
    """
    url = f"https://www.swiggy.com/dapi/restaurants/search/v3?lat={lat}&lng={lng}&str={requests.utils.quote(search_name)}&trackingId=undefined&submitAction=ENTER&queryUniqueId=undefined"
    try:
        r = requests.get(url, headers=HEADERS, timeout=12)
        if r.status_code == 200:
            data = r.json()
            # Recursive helper to find restaurant card
            found_card = [None]
            def extract_info(obj):
                if found_card[0] is not None:
                    return
                if isinstance(obj, dict):
                    if 'card' in obj and isinstance(obj['card'], dict) and 'card' in obj['card']:
                        info = obj['card']['card'].get('info', {})
                        if info:
                            # Match by ID or Name
                            name_match = search_name.lower() in info.get('name', '').lower()
                            id_match = (restaurant_id and str(info.get('id')) == str(restaurant_id))
                            if name_match or id_match:
                                found_card[0] = info
                                return
                    for k, v in obj.items():
                        extract_info(v)
                elif isinstance(obj, list):
                    for item in obj:
                        extract_info(item)
            extract_info(data)

            if found_card[0]:
                info = found_card[0]
                raw_rating = info.get('avgRating') or info.get('avgRatingString')
                raw_total = info.get('totalRatingsString')
                
                rating = float(raw_rating) if raw_rating and raw_rating != "NEW" else None
                
                # Parse total count like "46", "1.8K+", "100+"
                total_count = None
                if raw_total:
                    clean_str = str(raw_total).replace('+', '').replace('ratings', '').replace('rating', '').strip()
                    if 'K' in clean_str.upper() or 'k' in clean_str:
                        num = float(clean_str.upper().replace('K', '').strip())
                        total_count = int(num * 1000)
                    else:
                        digits = re.findall(r'\d+', clean_str)
                        if digits:
                            total_count = int("".join(digits))
                
                return {
                    "rating": rating,
                    "total_count": total_count,
                    "raw_rating": raw_rating,
                    "raw_total": raw_total
                }
    except Exception as e:
        print(f"[Swiggy] Error fetching {search_name}: {e}")
    return {"rating": None, "total_count": None}


def get_zomato_ratings(restaurant_slug):
    """
    Fetch dining and delivery ratings and review counts from Zomato.
    """
    url = f"https://www.zomato.com/hyderabad/{restaurant_slug}"
    try:
        r = requests.get(url, headers=HEADERS, timeout=12)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')
            
            # Check for __PRELOADED_STATE__
            match = re.search(r'window\.__PRELOADED_STATE__\s*=\s*JSON\.parse\("((?:\\.|[^"\\])*)"\);', r.text)
            if match:
                raw = match.group(1).encode('utf-8').decode('unicode-escape')
                data = json.loads(raw)
                
                # Look for SECTION_BASIC_INFO
                basic_info = None
                res_obj = data.get('pages', {}).get('restaurant', {})
                for k, v in res_obj.items():
                    if isinstance(v, dict) and 'sections' in v:
                        basic_info = v['sections'].get('SECTION_BASIC_INFO', {})
                        break

                if basic_info:
                    rating_new = basic_info.get('rating_new', {}).get('ratings', {})
                    dining_obj = rating_new.get('DINING', {})
                    delivery_obj = rating_new.get('DELIVERY', {})
                    
                    # Dining stats
                    dining_rating = float(dining_obj.get('rating')) if dining_obj.get('rating') else None
                    dining_count = int(dining_obj.get('reviewCount')) if dining_obj.get('reviewCount') and dining_obj.get('reviewCount').isdigit() else None
                    
                    # Delivery stats
                    delivery_rating = float(delivery_obj.get('rating')) if delivery_obj.get('rating') else None
                    delivery_count = int(delivery_obj.get('reviewCount')) if delivery_obj.get('reviewCount') and delivery_obj.get('reviewCount').isdigit() else None
                    
                    # Fallback aggregate
                    overall_rating = float(basic_info.get('rating', {}).get('aggregate_rating', 0)) or None
                    overall_votes = int(basic_info.get('rating', {}).get('votes', 0)) or None

                    return {
                        "dining_rating": dining_rating or overall_rating,
                        "dining_count": dining_count or overall_votes,
                        "delivery_rating": delivery_rating or overall_rating,
                        "delivery_count": delivery_count or overall_votes,
                        "overall_rating": overall_rating,
                        "overall_votes": overall_votes
                    }
                    
            # Fallback to LD+JSON
            scripts = soup.find_all('script', type='application/ld+json')
            for s in scripts:
                if s.string and 'aggregateRating' in s.string:
                    ld_data = json.loads(s.string)
                    agg = ld_data.get('aggregateRating', {})
                    return {
                        "dining_rating": float(agg.get('ratingValue', 0)),
                        "dining_count": int(agg.get('ratingCount', 0)),
                        "delivery_rating": float(agg.get('ratingValue', 0)),
                        "delivery_count": int(agg.get('ratingCount', 0)),
                    }
    except Exception as e:
        print(f"[Zomato] Error fetching {restaurant_slug}: {e}")
    return {"dining_rating": None, "dining_count": None, "delivery_rating": None, "delivery_count": None}


def get_google_ratings(place_name, place_id=None):
    """
    Fetch Google rating and review count.
    Supports SerpAPI (Google Maps / Reviews API), Google Places API, or configured metrics fallback.
    """
    serpapi_key = os.getenv("SERPAPI_KEY", SERPAPI_KEY).strip()
    google_places_key = os.getenv("GOOGLE_PLACES_API_KEY", GOOGLE_PLACES_API_KEY).strip()

    # 1. SerpAPI (Google Maps Reviews / Place API)
    if serpapi_key:
        try:
            target_desc = place_id if place_id else place_name
            print(f"[SerpAPI] Fetching live Google Maps ratings for '{place_name}' (ID: {place_id})...")
            params = {
                "engine": "google_maps",
                "hl": "en",
                "gl": "in",
                "api_key": serpapi_key
            }
            if place_id:
                params["place_id"] = place_id
            else:
                params["q"] = place_name

            r = requests.get("https://serpapi.com/search.json", params=params, timeout=15)
            if r.status_code == 200:
                data = r.json()
                
                # Check place_results or local_results
                pr = data.get("place_results")
                if not pr and "local_results" in data and len(data["local_results"]) > 0:
                    target_kw = "kipling" if "kipling" in place_name.lower() else ("casa" if "casa" in place_name.lower() else "")
                    for item in data["local_results"]:
                        if target_kw and target_kw in item.get("title", "").lower():
                            pr = item
                            break
                    if not pr:
                        pr = data["local_results"][0]

                if pr:
                    raw_rating = pr.get("rating")
                    raw_reviews = (
                        pr.get("reviews") 
                        or pr.get("reviews_count") 
                        or pr.get("user_ratings_total") 
                        or (pr.get("user_reviews") or {}).get("total_reviews")
                    )
                    
                    rating = float(raw_rating) if raw_rating is not None else None
                    
                    # Parse reviews count if formatted as int or string like "140", "140 reviews", "1.4k"
                    total_count = None
                    if raw_reviews is not None:
                        if isinstance(raw_reviews, int):
                            total_count = raw_reviews
                        else:
                            clean_str = str(raw_reviews).replace(",", "").replace("+", "").replace("reviews", "").strip()
                            if "k" in clean_str.lower():
                                num_part = float(re.findall(r"[\d\.]+", clean_str)[0])
                                total_count = int(num_part * 1000)
                            else:
                                digits = re.findall(r"\d+", clean_str)
                                if digits:
                                    total_count = int("".join(digits))

                    if rating is not None and total_count is not None:
                        print(f"[SerpAPI SUCCESS] {place_name} -> {total_count} reviews, {rating} stars")
                        return {"rating": rating, "total_count": total_count}
            else:
                print(f"[SerpAPI Warning] API returned status {r.status_code}: {r.text[:200]}")
        except Exception as e:
            print(f"[SerpAPI ERROR] Error fetching {place_name}: {e}")

    # 2. Google Places API if key provided
    if google_places_key:
        try:
            print(f"[Google Places API] Querying for '{place_name}'...")
            if not place_id:
                find_url = f"https://maps.googleapis.com/maps/api/place/findplacefromtext/json?input={requests.utils.quote(place_name)}&inputtype=textquery&fields=place_id,rating,user_ratings_total&key={google_places_key}"
                r = requests.get(find_url, timeout=10).json()
                if r.get('candidates'):
                    cand = r['candidates'][0]
                    return {
                        "rating": float(cand.get('rating')),
                        "total_count": int(cand.get('user_ratings_total'))
                    }
            else:
                det_url = f"https://maps.googleapis.com/maps/api/place/details/json?place_id={place_id}&fields=rating,user_ratings_total&key={google_places_key}"
                r = requests.get(det_url, timeout=10).json()
                if r.get('result'):
                    res = r['result']
                    return {
                        "rating": float(res.get('rating')),
                        "total_count": int(res.get('user_ratings_total'))
                    }
        except Exception as e:
            print(f"[Google Places API ERROR] Error for {place_name}: {e}")

    # 3. If no API key configured
    print(f"[Google Ratings] Note: Neither SERPAPI_KEY nor GOOGLE_PLACES_API_KEY is configured in .env.")
    print(f"[Google Ratings] To fetch live Google ratings & review counts, add your SERPAPI_KEY to .env.")
    return {"rating": None, "total_count": None}


def fetch_all_ratings():
    """
    Fetches all metrics required for the daily spreadsheet:
    - Kiplings: Google (Count, Rating), Zomato (Count, Rating)
    - Casa Loco Express: Google (Count, Rating), Zomato - Delivery (Count, Rating), Swiggy - Delivery (Count, Rating)
    """
    print("\n==========================================")
    print("Fetching today's restaurant ratings...")
    print("==========================================")

    # 1. Kiplings
    k_cfg = RESTAURANTS_CONFIG["kiplings"]
    print(f"Fetching data for {k_cfg['display_name']}...")
    k_zomato = get_zomato_ratings(k_cfg["zomato_slug"])
    k_google = get_google_ratings(k_cfg["google_query"], k_cfg["google_place_id"])

    # 2. Casa Loco Express
    c_cfg = RESTAURANTS_CONFIG["casa_loco"]
    print(f"Fetching data for {c_cfg['display_name']}...")
    c_swiggy = get_swiggy_ratings(c_cfg["swiggy_name"], c_cfg["swiggy_id"], c_cfg["lat"], c_cfg["lng"])
    c_zomato = get_zomato_ratings(c_cfg["zomato_slug"])
    c_google = get_google_ratings(c_cfg["google_query"], c_cfg["google_place_id"])

    data = {
        "kiplings": {
            "google_count": k_google.get("total_count"),
            "google_rating": k_google.get("rating"),
            "zomato_count": k_zomato.get("dining_count") or k_zomato.get("overall_votes"),
            "zomato_rating": k_zomato.get("dining_rating") or k_zomato.get("overall_rating"),
        },
        "casa_loco": {
            "google_count": c_google.get("total_count"),
            "google_rating": c_google.get("rating"),
            "zomato_count": c_zomato.get("delivery_count") or c_zomato.get("overall_votes"),
            "zomato_rating": c_zomato.get("delivery_rating") or c_zomato.get("overall_rating"),
            "swiggy_count": c_swiggy.get("total_count"),
            "swiggy_rating": c_swiggy.get("rating"),
        }
    }

    print("\nSummary of Fetched Ratings:")
    k_g_c = data['kiplings']['google_count'] if data['kiplings']['google_count'] is not None else "N/A"
    k_g_r = data['kiplings']['google_rating'] if data['kiplings']['google_rating'] is not None else "N/A"
    c_g_c = data['casa_loco']['google_count'] if data['casa_loco']['google_count'] is not None else "N/A"
    c_g_r = data['casa_loco']['google_rating'] if data['casa_loco']['google_rating'] is not None else "N/A"

    print(f"- Kiplings -> Google: {k_g_c} reviews ({k_g_r} stars), Zomato: {data['kiplings']['zomato_count']} reviews ({data['kiplings']['zomato_rating']} stars)")
    print(f"- Casa Loco -> Google: {c_g_c} reviews ({c_g_r} stars), Zomato Delivery: {data['casa_loco']['zomato_count']} reviews ({data['casa_loco']['zomato_rating']} stars), Swiggy Delivery: {data['casa_loco']['swiggy_count']} reviews ({data['casa_loco']['swiggy_rating']} stars)")
    print("==========================================\n")

    return data


if __name__ == "__main__":
    res = fetch_all_ratings()
    print(json.dumps(res, indent=2))
