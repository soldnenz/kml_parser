import json
import re
import simplekml
import math

def parse_dms(dms_str):
    """
    Parse a DMS string into decimal degrees.
    Examples: 'N433604', 'E0713936', etc.
    """
    # Clean up the input, removing any spaces
    dms_str = dms_str.strip()
    
    # Extract direction (N, S, E, W)
    direction = dms_str[0]
    
    # Extract numeric part and ensure it's a string
    numeric_part = dms_str[1:].strip()
    
    # Check if this is latitude (DDMMSS) or longitude (DDDMMSS)
    if direction in ['E', 'W']:
        # Longitude format: DDDMMSS
        if len(numeric_part) == 7:
            d = int(numeric_part[0:3])
            m = int(numeric_part[3:5])
            s = int(numeric_part[5:7])
        else:
            # Handle shorter formats by padding with zeros
            padded = numeric_part.zfill(7)
            d = int(padded[0:3])
            m = int(padded[3:5])
            s = int(padded[5:7])
    else:
        # Latitude format: DDMMSS
        if len(numeric_part) == 6:
            d = int(numeric_part[0:2])
            m = int(numeric_part[2:4])
            s = int(numeric_part[4:6])
        else:
            # Handle shorter formats by padding with zeros
            padded = numeric_part.zfill(6)
            d = int(padded[0:2])
            m = int(padded[2:4])
            s = int(padded[4:6])
    
    # Convert to decimal degrees
    dec = d + m/60 + s/3600
    
    # Apply sign based on direction
    if direction in ['S', 'W']:
        dec = -dec
    
    return dec

def create_circle_polygon(center_lat, center_lon, radius_meters, num_points=36):
    """
    Create a circular polygon based on center and radius.
    Returns a list of (lon, lat) tuples for KML.
    """
    # Constants for Earth radius
    EARTH_RADIUS = 6378137  # meters at equator
    
    # Convert radius from meters to degrees (approximate)
    radius_deg_lat = (radius_meters / EARTH_RADIUS) * (180 / math.pi)
    radius_deg_lon = radius_deg_lat / math.cos(math.radians(center_lat))
    
    # Generate points around the circle
    points = []
    for i in range(num_points + 1):  # +1 to close the circle
        angle = 2 * math.pi * i / num_points
        lat = center_lat + radius_deg_lat * math.sin(angle)
        lon = center_lon + radius_deg_lon * math.cos(angle)
        points.append((lon, lat))  # KML expects (lon, lat)
    
    return points

def extract_radius(radius_str):
    """
    Extract radius value in meters from string like "R=5000 м"
    """
    match = re.search(r'R[=-](\d+)', radius_str)
    if match:
        return int(match.group(1))
    return 5000  # Default radius if not found

def get_coords(coord_str):
    """
    Parse coordinates from the zone definition string.
    Handles both polygon and circle formats.
    """
    # Replace Cyrillic 'Е' with Latin 'E' if present
    coord_str = coord_str.replace('Е', 'E')
    
    # Remove extra spaces for easier parsing
    coord_str = re.sub(r'\s+', ' ', coord_str.strip())
    
    # Check for circle format
    circle_format = False
    radius = 5000  # default radius in meters
    
    if "R=" in coord_str or "R-" in coord_str:
        circle_format = True
        radius = extract_radius(coord_str)
    
    # Check for the reversed format pattern (e.g., "460755N 0805610E")
    reversed_format = re.findall(r'(\d{6})([NS])[ \-]?(\d{7})([EW])', coord_str)
    if reversed_format:
        if circle_format:
            # Only use the first coordinate pair as center for circle
            lat_val, lat_dir, lon_val, lon_dir = reversed_format[0]
            lat = parse_dms(f"{lat_dir}{lat_val}")
            lon = parse_dms(f"{lon_dir}{lon_val}")
            return create_circle_polygon(lat, lon, radius)
        else:
            # Process as polygon
            coords = []
            for lat_val, lat_dir, lon_val, lon_dir in reversed_format:
                lat = parse_dms(f"{lat_dir}{lat_val}")
                lon = parse_dms(f"{lon_dir}{lon_val}")
                coords.append((lon, lat))  # KML expects (lon, lat)
            
            # Make sure the polygon is closed
            if coords and coords[0] != coords[-1]:
                coords.append(coords[0])
                
            return coords
    
    # Standard format (e.g., "N433604 E0765618")
    pairs = re.findall(r'([NS])[ \-]?(\d{6})[ \-,]*([EW])[ \-]?(\d{7})', coord_str)
    
    if pairs:
        if circle_format:
            # Only use the first coordinate pair as center for circle
            lat_dir, lat_val, lon_dir, lon_val = pairs[0]
            lat = parse_dms(f"{lat_dir}{lat_val}")
            lon = parse_dms(f"{lon_dir}{lon_val}")
            return create_circle_polygon(lat, lon, radius)
        else:
            # Process as polygon
            coords = []
            for lat_dir, lat_val, lon_dir, lon_val in pairs:
                lat = parse_dms(f"{lat_dir}{lat_val}")
                lon = parse_dms(f"{lon_dir}{lon_val}")
                coords.append((lon, lat))  # KML expects (lon, lat)
            
            # Make sure the polygon is closed
            if coords and coords[0] != coords[-1]:
                coords.append(coords[0])
                
            return coords
    
    return []

def main():
    # Load zones data
    with open('zones.json', 'r', encoding='utf-8') as f:
        zones = json.load(f)

    # Create KML document
    kml = simplekml.Kml()

    # Define single style for all zones - red with thin border
    zone_style = simplekml.Style()
    zone_style.linestyle.color = simplekml.Color.red
    zone_style.linestyle.width = 1.0  # Thinner border
    zone_style.polystyle.color = simplekml.Color.changealphaint(50, simplekml.Color.red)  # Red fill with 50% opacity

    # Process each zone
    successful_zones = 0
    failed_zones = 0
    
    for zone in zones:
        # Skip excluded zones
        if "Исключена приказом" in zone.get('1', ''):
            continue
            
        name = zone.get('1', '')
        coord_str = zone.get('2', '')
        alt_range = zone.get('3', '')
        alt_limit = zone.get('4', '')
        schedule = zone.get('5', '')

        if not name or not coord_str:
            continue
            
        # Get coordinates - KML expects (lon, lat) pairs
        coords = get_coords(coord_str)
        if not coords:
            print(f"Failed to parse coordinates for zone {name}: {coord_str}")
            failed_zones += 1
            continue

        # Set description with zone information
        description = f"Altitude: {alt_range}\nLimit: {alt_limit}\nSchedule: {schedule}"

        # Create polygon with the appropriate style
        poly = kml.newpolygon(name=name, outerboundaryis=coords, description=description)
        poly.style = zone_style
        
        successful_zones += 1

    # Save to file
    kml.save("zones.kml")
    print(f"KML file generated successfully - {successful_zones} zones processed, {failed_zones} zones failed")

if __name__ == "__main__":
    main()

# Usage:
# Place this script alongside your 'zones.json' file and run:
# python create_kml_zones.py
