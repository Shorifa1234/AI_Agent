# Updated category links with all URLs
CATEGORY_LINKS = {
    'Coffee & Cocktail Tables': ['https://www.fdbmobler.com/collections/coffee-tables'],
    'Side & End Tables': ['https://www.fdbmobler.com/collections/side-tables'],
    'Dining Tables': ['https://www.fdbmobler.com/collections/dining-tables'],
    'Consoles': ['https://www.fdbmobler.com/collections/sideboard'],
    'Desks': ['https://www.fdbmobler.com/collections/desks'],
    'Bookcases': [
        'https://www.fdbmobler.com/collections/shelves',
        'https://www.fdbmobler.com/collections/book-cases',
        'https://www.fdbmobler.com/collections/shelves-1'
    ],
    'Cabinets': ['https://www.fdbmobler.com/collections/display-cabinets'],
    'Dining Chairs': ['https://www.fdbmobler.com/collections/dining-table-chairs-1'],
    'Bar Stools': [
        'https://www.fdbmobler.com/collections/bar-stools',
        'https://www.fdbmobler.com/collections/stools'
    ],
    'Sofas & Loveseats': [
        'https://www.fdbmobler.com/collections/module-sofas',
        'https://www.fdbmobler.com/collections/2-person-sofa',
        'https://www.fdbmobler.com/collections/2-5-person-sofa',
        'https://www.fdbmobler.com/collections/3-person-sofa',
        'https://www.fdbmobler.com/collections/cushions-for-sofas'
    ],
    'Lounge Chairs': [
        'https://www.fdbmobler.com/collections/armchairs',
        'https://www.fdbmobler.com/collections/cushions-for-easy-chairs'
    ],
    'Benches': ['https://www.fdbmobler.com/collections/benches'],
    'Pendants': ['https://www.fdbmobler.com/collections/pedants'],
    'Sconces': ['https://www.fdbmobler.com/collections/wall-lamps'],
    'Table Lamps': ['https://www.fdbmobler.com/collections/table-lamps'],
    'Floor Lamps': ['https://www.fdbmobler.com/collections/floor-lamps'],
    'Mirrors': ['https://www.fdbmobler.com/collections/mirrors'],
    'Pillows & Throws': ['https://www.fdbmobler.com/collections/cushions'],
    'Vases': ['https://www.fdbmobler.com/collections/vases'],
    'Boxes': [
        'https://www.fdbmobler.com/collections/boxes',
        'https://www.fdbmobler.com/collections/apple-boxes'
    ],
    'Outdoor Seating': [
        'https://www.fdbmobler.com/collections/garden-chairs',
        'https://www.fdbmobler.com/collections/lounge-furniture',
        'https://www.fdbmobler.com/collections/garden-benches'
    ],
    'Outdoor Tables': ['https://www.fdbmobler.com/collections/garden-tables'],
    'Outdoor Accessories': ['https://www.fdbmobler.com/collections/accessories-for-the-garden'],
}

# Calculate total
total_links = sum(len(links) for links in CATEGORY_LINKS.values())
print(f"Total categories: {len(CATEGORY_LINKS)}")
print(f"Total category links: {total_links}")
print(f"\nCategories with multiple links:")
for cat, links in CATEGORY_LINKS.items():
    if len(links) > 1:
        print(f"  {cat}: {len(links)} links")
