from pathlib import Path
from datetime import datetime, timedelta
import xlsxwriter

OUTPUT = Path("build-output")
OUTPUT.mkdir(exist_ok=True)

# Running this script again overwrites files with these names.
# Keep manually edited/tested copies somewhere else.

SHEETS = [
    "Dashboard", "Setup", "Save the Date", "Theme Board", "Calendar",
    "Planning Timeline", "Wedding Itinerary", "Packing List",
    "Vendor Choices", "Venue Options", "Wedding Budget",
    "Contact Directory", "Guest List", "Seating Plan", "Wedding Party",
    "Food and Drinks", "Photoshoot", "Photo Gallery", "Music",
    "Gifts and Thank You", "Honeymoon Budget",
]

CATEGORIES = [
    "Venue", "Catering", "Photography", "Attire", "Flowers",
    "Music", "Stationery", "Transport", "Ceremony", "Other",
]

CHOICES = {
    "Categories": CATEGORIES,
    "TaskStatus": ["To do", "In progress", "Done"],
    "RSVP": ["Awaiting reply", "Attending", "Declined"],
    "Meals": ["Standard", "Vegetarian", "Vegan", "Child", "Other"],
    "Delivery": ["Email", "Printed", "Message"],
    "InviteStatus": ["To prepare", "Ready", "Sent"],
    "PackingStatus": ["To pack", "Packed"],
    "Bags": ["Wedding day", "Overnight", "Emergency kit", "Honeymoon"],
    "VendorStatus": ["Considering", "Contacted", "Shortlisted", "Chosen", "Rejected"],
    "VenueStatus": ["Considering", "Viewing booked", "Visited", "Favourite",
                    "Booked", "Rejected"],
    "Party": ["Partner one", "Partner two", "Both"],
    "Roles": ["Person of honour", "Best person", "Attendant",
              "Flower person", "Ring bearer", "Other"],
    "Courses": ["Starter", "Main", "Dessert", "Drinks", "Other"],
    "MenuStatus": ["Considering", "Selected", "Rejected"],
    "ShotCategory": ["Couple", "Family", "Wedding party", "Details", "Other"],
    "Priority": ["Must have", "Nice to have"],
    "ShotStatus": ["Planned", "Captured"],
    "MusicMoments": ["Walking in", "Signing", "Walking out", "First dance",
                    "Dinner", "Party", "Do not play", "Other"],
    "Thanks": ["To write", "Written", "Sent"],
    "HoneyCategories": ["Accommodation", "Flights", "Transport", "Food",
                        "Experiences", "Insurance", "Other"],
    "TableNames": [str(i) for i in range(1, 21)],
}

D = lambda y, m, d: datetime(y, m, d)

CLIENTS = [
    {
        "file": "Wedding-Planner.xlsx",
        "clean": True,
        "one": "", "two": "", "date": None,
        "venue": "", "location": "", "budget": None, "honeymoon": None,
    },
    {
        "file": "Client-A.xlsx",
        "clean": False,
        "one": "David", "two": "Pia",
        "date": D(2027, 10, 10),
        "venue": "Pringle Rock Distillery",
        "location": "Cape Town, South Africa",
        "budget": 180000, "honeymoon": 50000,
    },
    {
        "file": "Client-B.xlsx",
        "clean": False,
        "one": "Eben", "two": "Sharks",
        "date": D(2027, 11, 6),
        "venue": "The Shark Tank",
        "location": "Durban, South Africa",
        "budget": 120000, "honeymoon": 35000,
    },
]
THEME_PALETTE = [
    ("Sage", "#81917A"),
    ("Blush", "#E9D1C8"),
    ("Ivory", "#FAF8F2"),
    ("Forest", "#344A3B"),
    ("Eucalyptus", "#A6B49F"),
    ("Olive", "#808000"),
    ("Champagne", "#F7E7CE"),
    ("Sand", "#D8C3A5"),
    ("Terracotta", "#C4704F"),
    ("Dusty rose", "#C9959C"),
    ("Burgundy", "#800020"),
    ("Lavender", "#B8A2C8"),
    ("Dusty blue", "#8FA9BF"),
    ("Navy", "#203864"),
    ("Gold", "#D4AF37"),
    ("Charcoal", "#36454F"),
    ("White", "#FFFFFF"),
    ("Black", "#000000"),
]

def build(client):
    wb = xlsxwriter.Workbook(OUTPUT / client["file"])
    wb.set_calc_mode("auto")
    wb.set_properties({
        "title": "Wedding Planner — assessment adaptation",
        "comments": (
            "Assessment adaptation of Johan Botha's supplied reference pack. "
            "No commercial rights to supplied materials are implied."
        ),
    })

    sage = "#81917A"
    dark = "#344A3B"
    ivory = "#FAF8F2"
    amber = "#FFF3D6"
    pale = "#EDF2E9"
    blush = "#E9D1C8"

    f = {
        "calc_int": wb.add_format({
            "bg_color": pale,
            "num_format": "#,##0",
            "font_color": dark,
        }),        
        "title": wb.add_format({
            "font_name": "Georgia", "font_size": 24,
            "font_color": dark, "bold": True,
        }),
        "text": wb.add_format({
            "font_name": "Calibri", "font_size": 11,
            "font_color": dark, "valign": "top",
        }),
        "wrap": wb.add_format({
            "font_color": dark, "text_wrap": True, "valign": "top",
        }),
        "input": wb.add_format({"bg_color": amber, "font_color": dark}),
        "date": wb.add_format({
            "bg_color": amber, "num_format": "dd mmm yyyy",
            "font_color": dark,
        }),
        "money": wb.add_format({
            "bg_color": amber, "num_format": "#,##0.00",
            "font_color": dark,
        }),
        "integer": wb.add_format({
            "bg_color": amber, "num_format": "0", "font_color": dark,
        }),
        "time": wb.add_format({
            "bg_color": amber, "num_format": "hh:mm", "font_color": dark,
        }),
        "calc": wb.add_format({
            "bg_color": pale, "num_format": "#,##0.00",
            "font_color": dark,
        }),
        "calc_date": wb.add_format({
            "bg_color": pale, "num_format": "dd mmm yyyy",
            "font_color": dark,
        }),
        "calc_text": wb.add_format({"bg_color": pale, "font_color": dark}),
        "percent": wb.add_format({
            "bg_color": pale, "num_format": "0%", "font_color": dark,
        }),
        "section": wb.add_format({
            "bg_color": sage, "font_color": "white", "bold": True,
        }),
        "link": wb.add_format({
            "font_color": dark, "underline": True,
        }),
        "warning": wb.add_format({
            "bg_color": "#F7DAD5", "font_color": "#8B2525",
        }),
        "good": wb.add_format({"bg_color": "#DDEBD7"}),
        "photo": wb.add_format({
            "bg_color": blush, "align": "center",
            "valign": "vcenter", "font_color": dark, "border": 1,
        }),
    }

    start = wb.add_worksheet("Start Here")
    ws = {name: wb.add_worksheet(name) for name in SHEETS}
    lists = wb.add_worksheet("Lists")

    all_sheets = {"Start Here": start, **ws}
    for name, sh in all_sheets.items():
        sh.hide_gridlines(2)
        sh.set_tab_color(sage)
        sh.set_default_row(21)
        sh.set_row(0, 36)        
        sh.set_column("A:A", 3)
        sh.set_column("B:B", 30)
        sh.set_column("C:D", 22)
        sh.set_column("E:L", 19)
        sh.set_landscape()
        sh.set_paper(9)  # A4
        sh.fit_to_pages(1, 0)
        sh.set_margins(0.3, 0.3, 0.4, 0.4)
        sh.set_footer("&LAssessment adaptation — Johan Botha reference pack&R&P / &N")
        sh.write("B1", name, f["title"])
        if name != "Start Here":
            sh.write_url("B2", "internal:'Start Here'!B1",
                         f["link"], "← Start Here / contents")
        sh.write_formula(
            "B3",
            '=IF(AND(PartnerOne="",PartnerTwo=""),'
            '"Your wedding planner",PartnerOne&" & "&PartnerTwo)',
            f["text"],
        )
        sh.write_formula(
            "B4",
            '="Amounts: "&Currency&" | Amber = input; green = calculated"',
            f["text"],
        )
        if name not in ("Dashboard","Setup","Theme Board","Photo Gallery","Start Here","Calendar"):
            sh.freeze_panes(8, 0)
        sh.repeat_rows(0, 7)

    # Dropdown sources: internal only, no external workbook references.
    for col, (name, values) in enumerate(CHOICES.items()):
        lists.write(0, col, name)
        lists.write_column(1, col, values)
        letter = xlsxwriter.utility.xl_col_to_name(col)
        wb.define_name(name, f"=Lists!${letter}$2:${letter}${len(values)+1}")

    # Theme palette lookup table, placed after the existing dropdown lists.
    palette_name_col = len(CHOICES)
    palette_hex_col = palette_name_col + 1

    lists.write(0, palette_name_col, "Theme colour")
    lists.write(0, palette_hex_col, "Hex")

    for row, (name, hex_code) in enumerate(THEME_PALETTE, start=1):
        lists.write_string(row, palette_name_col, name)
        lists.write_string(row, palette_hex_col, hex_code)

    name_letter = xlsxwriter.utility.xl_col_to_name(palette_name_col)
    hex_letter = xlsxwriter.utility.xl_col_to_name(palette_hex_col)
    last_palette_row = len(THEME_PALETTE) + 1

    wb.define_name(
        "ThemeColourNames",
        f"='Lists'!${name_letter}$2:"
        f"${name_letter}${last_palette_row}",
    )
    wb.define_name(
        "ThemeColourHex",
        f"='Lists'!${hex_letter}$2:"
        f"${hex_letter}${last_palette_row}",
    )
    
    lists.hide()

    # Central configuration / named cells.
    setup = ws["Setup"]
    fields = [
        ("Partner one", client["one"], "PartnerOne", "input"),
        ("Partner two", client["two"], "PartnerTwo", "input"),
        ("Wedding date", client["date"], "WeddingDate", "date"),
        ("Venue", client["venue"], "Venue", "input"),
        ("Location", client["location"], "Location", "input"),
        ("Wedding budget", client["budget"], "WeddingBudget", "money"),
        ("Currency code / label", "ZAR", "Currency", "input"),
        ("Honeymoon budget", client["honeymoon"], "HoneymoonBudget", "money"),
        ("Palette colour 1", "#81917A", "ColourOne", "input"),
        ("Palette colour 2", "#E9D1C8", "ColourTwo", "input"),
        ("Palette colour 3", "#FAF8F2", "ColourThree", "input"),
    ]
    for row, (label, value, name, style) in enumerate(fields, start=7):
        setup.write(row, 1, label, f["text"])
        setup.write(row, 3, value, f[style])
        wb.define_name(name, f"='Setup'!$D${row+1}")

    setup.data_validation("D10", {
        "validate": "date", "criteria": "between",
        "minimum": D(2020, 1, 1), "maximum": D(2100, 12, 31),
    })
    for cell in ["D13", "D15"]:
        setup.data_validation(cell, {
            "validate": "decimal", "criteria": ">=", "value": 0,
        })
    setup.merge_range(
        "B21:H23",
        "Use one currency per workbook. Changing the currency label does not "
        "convert amounts. Choose palette colours using the dropdowns on "
        "Theme Board. Swatches and hex codes update automatically from the "
        "predefined palette. No macros, account or cloud connection required.",
        f["wrap"],
    )

    # Generic native Excel Table builder.
    def table(sheet, name, headers, rows, styles=None,
              formulas=None, dropdowns=None, nonnegative=None):
        sh = ws[sheet]
        styles = styles or {}
        formulas = formulas or {}
        dropdowns = dropdowns or {}
        nonnegative = nonnegative or []
        count = max(30, len(rows) + 5)
        data = [
            list(row) + [None] * (len(headers) - len(row))
            for row in rows
        ]
        data += [[None] * len(headers) for _ in range(count - len(data))]

        columns = []
        for header in headers:
            spec = {
                "header": header,
                "format": f[styles.get(header, "input")],
            }
            if header in formulas:
                spec["formula"] = formulas[header]
                spec["format"] = f[
                    "calc_date" if header == "Due date" else
                    "calc_text" if header == "Payment status" else "calc"
                ]
            columns.append(spec)

        sh.write(
            "B6",
            "Enter records in the table. Use Tab from its last cell to add a row.",
            f["text"],
        )
        sh.add_table(7, 1, 7 + count, len(headers), {
            "name": name, "style": "Table Style Medium 4",
            "columns": columns, "data": data,
        })

        # Validation is deliberately pre-applied through worksheet row 5000.
        for header, source in dropdowns.items():
            col = 1 + headers.index(header)
            sh.data_validation(8, col, 4999, col, {
                "validate": "list", "source": "=" + source,
                "error_type": "stop",
                "error_title": "Choose a listed value",
                "error_message": "Use the dropdown list for this field.",
            })
        for header in nonnegative:
            col = 1 + headers.index(header)
            sh.data_validation(8, col, 4999, col, {
                "validate": "decimal", "criteria": ">=", "value": 0,
                "error_type": "stop",
                "error_message": "Enter a number of zero or greater.",
            })
        for header, style in styles.items():
            if style == "date":
                col = 1 + headers.index(header)
                sh.data_validation(8, col, 4999, col, {
                    "validate": "date", "criteria": "between",
                    "minimum": D(2020, 1, 1),
                    "maximum": D(2100, 12, 31),
                })

        sh.print_area(0, 1, max(20, 8 + len(rows)), len(headers))
        return sh

    demo = not client["clean"]
    wedding = client["date"] or D(2027, 9, 18)
    one = client["one"] or "Partner one"
    two = client["two"] or "Partner two"
    scale = 1 if client["file"] != "Client-B.xlsx" else 0.65
    amount = lambda x: round(x * scale, 2)
    sample = lambda rows: rows if demo else []

    table("Save the Date", "Invites",
          ["Household", "Email", "Delivery", "Send by", "Status"],
          sample([
              ["Morgan household", "morgan@example.test", "Email",
               wedding - timedelta(days=180), "Ready"],
              ["Lee household", "lee@example.test", "Printed",
               wedding - timedelta(days=180), "To prepare"],
          ]),
          styles={"Send by": "date"},
          dropdowns={"Delivery": "Delivery", "Status": "InviteStatus"})

    table("Planning Timeline", "Tasks",
          ["Task", "Category", "Owner", "Days before", "Fixed date",
           "Due date", "Status", "Notes"],
          sample([
              ["Book venue", "Venue", one, 240, None, None, "Done", ""],
              ["Send save-the-dates", "Stationery", two, 180, None,
               None, "In progress", ""],
              ["Confirm menu", "Catering", one, 60, None, None, "To do", ""],
              ["Final guest numbers", "Other", two, 21, None,
               None, "To do", ""],
              ["Pack overnight bag", "Other", one, 2, None,
               None, "To do", ""],
          ]),
          styles={"Days before": "integer", "Fixed date": "date"},
          formulas={
              "Due date":
              '=IF([@Task]="","",IF([@[Fixed date]]<>"",[@[Fixed date]],'
              'IF(AND(WeddingDate<>"",ISNUMBER([@[Days before]])),'
              'WeddingDate-[@[Days before]],"")))'
          },
          dropdowns={"Category": "Categories", "Status": "TaskStatus"},
          nonnegative=["Days before"])

    itinerary = table(
        "Wedding Itinerary", "Itinerary",
        ["Time", "Moment", "Location", "Responsible", "Notes"],
        sample([
            [9/24, "Getting ready", "Suite", one, "Allow time for breakfast"],
            [14/24, "First look", "Garden", "Photographer", ""],
            [15/24, "Ceremony", "Ceremony space", "Celebrant", ""],
            [16/24, "Photos and refreshments", "Terrace", two, ""],
            [18/24, "Dinner and speeches", "Reception", "MC", ""],
            [20/24, "First dance", "Dance floor", "DJ", ""],
        ]),
        styles={"Time": "time"},
    )
    itinerary.data_validation("B9:B5000", {
        "validate": "time", "criteria": "between",
        "minimum": 0, "maximum": 0.999999,
    })
    itinerary.write("H6", "Sort Time ascending before printing.", f["text"])

    table("Packing List", "Packing",
          ["Item", "Bag", "Quantity", "Packed by", "Status"],
          sample([
              ["Rings", "Wedding day", 2, two, "To pack"],
              ["Comfortable shoes", "Overnight", 2, one, "To pack"],
              ["Sewing kit", "Emergency kit", 1, one, "Packed"],
          ]),
          styles={"Quantity": "integer"},
          dropdowns={"Bag": "Bags", "Status": "PackingStatus"},
          nonnegative=["Quantity"])

    table("Vendor Choices", "Vendors",
          ["Supplier", "Category", "Contact", "Quote", "Status", "Notes"],
          sample([
              ["Daylight Studio", "Photography", "photo@example.test",
               amount(18000), "Chosen", "Eight hours"],
              ["Meadow Flowers", "Flowers", "flowers@example.test",
               amount(12500), "Shortlisted", "Seasonal flowers"],
          ]),
          styles={"Quote": "money"},
          dropdowns={"Category": "Categories", "Status": "VendorStatus"},
          nonnegative=["Quote"])

    table("Venue Options", "Venues",
          ["Venue", "Location", "Capacity", "Quote", "Contact",
           "Viewing date", "Status", "Notes"],
          sample([
              [client["venue"], client["location"], 100, amount(48000),
               "venue@example.test", wedding - timedelta(days=250),
               "Booked", "Tables and chairs included"],
              ["Mountain Lodge", "Winelands", 80, amount(55000),
               "lodge@example.test", wedding - timedelta(days=245),
               "Visited", "Accommodation available"],
          ]),
          styles={"Capacity": "integer", "Quote": "money", "Viewing date": "date"},
          dropdowns={"Status": "VenueStatus"},
          nonnegative=["Capacity", "Quote"])

    budget_formulas = {
        "Outstanding":
            '=IF([@Expense]="","",IF([@[Planned cost]]="","",'
            'MAX(0,[@[Planned cost]]-[@[Paid to date]])))',
        "Payment status":
            '=IF([@Expense]="","",IF([@[Planned cost]]="","Cost needed",'
            'IF([@[Paid to date]]>[@[Planned cost]],"Overpaid",'
            'IF([@Outstanding]=0,"Paid",'
            'IF([@[Paid to date]]>0,"Part paid","Unpaid")))))',
    }
    budget_headers = [
        "Expense", "Category", "Planned cost", "Paid to date",
        "Outstanding", "Payment due", "Payment status", "Notes",
    ]
    budget_styles = {
        "Planned cost": "money", "Paid to date": "money",
        "Payment due": "date",
    }

    budget_rows = [
        [client["venue"], "Venue", amount(48000), amount(24000),
         None, wedding - timedelta(days=60), None, "Booked venue"],
        ["Catering package", "Catering", amount(42000), amount(10000),
         None, wedding - timedelta(days=30), None,
         "Includes service; menu costing is a separate estimate"],
        ["Daylight Studio", "Photography", amount(18000), amount(6000),
         None, wedding - timedelta(days=30), None, ""],
        ["Outfits", "Attire", amount(15000), amount(7500),
         None, wedding - timedelta(days=45), None, ""],
        ["Flowers", "Flowers", amount(12500), 0,
         None, wedding - timedelta(days=20), None, ""],
        ["DJ", "Music", amount(8500), 0,
         None, wedding - timedelta(days=14), None, ""],
        ["Stationery", "Stationery", amount(3200), amount(3200),
         None, wedding - timedelta(days=90), None, ""],
        ["Contingency", "Other", amount(5000), 0,
         None, None, None, "Allowance, not a supplier invoice"],
    ]

    for sheet, name, rows, category_source in [
        ("Wedding Budget", "Budget", sample(budget_rows), "Categories"),
        ("Honeymoon Budget", "Honeymoon", sample([
            ["Accommodation", "Accommodation", amount(24000), amount(8000),
             None, wedding - timedelta(days=30), None, ""],
            ["Flights", "Flights", amount(18000), amount(6000),
             None, wedding - timedelta(days=60), None, ""],
            ["Boat trip", "Experiences", amount(2800), 0,
             None, wedding + timedelta(days=3), None, ""],
        ]), "HoneyCategories"),
    ]:
        sh = table(sheet, name, budget_headers, rows,
                   styles=budget_styles, formulas=budget_formulas,
                   dropdowns={"Category": category_source},
                   nonnegative=["Planned cost", "Paid to date"])
        sh.conditional_format("F9:F5000", {
            "type": "cell", "criteria": "<", "value": 0,
            "format": f["warning"],
        })
        sh.conditional_format("H9:H5000", {
            "type": "formula",
            "criteria": '=AND(ISNUMBER(G9),G9<TODAY(),F9>0)',
            "format": f["warning"],
        })
        sh.write("K8", "Planned", f["section"])
        sh.write_formula("L8", f"=SUM({name}[Planned cost])", f["calc"])
        sh.write("K9", "Paid", f["section"])
        sh.write_formula("L9", f"=SUM({name}[Paid to date])", f["calc"])
        sh.write("K10", "Outstanding", f["section"])
        sh.write_formula("L10", f"=SUM({name}[Outstanding])", f["calc"])
        sh.write("K11", "Remaining budget", f["section"])
        limit = "WeddingBudget" if name == "Budget" else "HoneymoonBudget"
        sh.write_formula(
            "L11", f'=IF({limit}="","",{limit}-SUM({name}[Planned cost]))',
            f["calc"],
        )
        sh.print_area(0, 1, max(20, 8 + len(rows)), 11)

    table("Contact Directory", "Contacts",
          ["Name", "Role", "Phone", "Email", "Notes"],
          sample([
              ["Casey Morgan", "Coordinator", "+27 00 000 0000",
               "coordinator@example.test", "Fictional demonstration contact"],
              ["Robin Lee", "Photographer", "+27 00 000 0000",
               "photo@example.test", ""],
          ]))

    guests = []
    if demo:
        for i, name in enumerate([
            "Casey Morgan", "Robin Lee", "Jamie Patel", "Avery Smith",
            "Riley Brown", "Drew Williams", "Sky Jacobs", "Quinn Davis",
            "Reese Martin", "Blair Wilson", "Cameron Botha", "Charlie Green",
        ]):
            status = "Attending" if i < 8 else (
                "Awaiting reply" if i < 11 else "Declined"
            )
            guests.append([
                f"G{i+1:03}", name, f"H{i//2+1:02}",
                "Partner one" if i % 2 else "Partner two",
                f"guest{i+1}@example.test", status,
                "Vegetarian" if i % 3 == 0 else "Standard",
                "Nuts" if i == 2 else "",
                str(1 if i < 4 else 2) if i < 7 else "",
                "",
            ])

    table("Guest List", "Guests",
          ["Guest ID", "Full name", "Household", "Party", "Email", "RSVP",
           "Meal", "Allergies", "Table", "Notes"],
          guests,
          dropdowns={"Party": "Party", "RSVP": "RSVP",
                     "Meal": "Meals", "Table": "TableNames"})

    table("Wedding Party", "WeddingParty",
          ["Name", "Role", "Contact", "Responsibility", "Notes"],
          sample([
              ["Casey Morgan", "Person of honour", "casey@example.test",
               "Keep rings safe", ""],
              ["Robin Lee", "Best person", "robin@example.test",
               "Welcome guests", ""],
          ]),
          dropdowns={"Role": "Roles"})

    table("Food and Drinks", "Menu",
          ["Menu item", "Course", "Unit cost", "Quantity",
           "Line total", "Dietary notes", "Status"],
          sample([
              ["Garden salad", "Starter", 95, 8, None, "Vegetarian", "Selected"],
              ["Roast chicken", "Main", 220, 5, None, "", "Selected"],
              ["Mushroom risotto", "Main", 180, 3, None,
               "Vegetarian", "Selected"],
              ["Lemon dessert", "Dessert", 85, 8, None, "", "Considering"],
              ["Soft drinks", "Drinks", 30, 16, None, "", "Selected"],
          ]),
          styles={"Unit cost": "money", "Quantity": "integer"},
          formulas={
              "Line total":
                  '=IF([@[Menu item]]="","",'
                  'IF(OR([@[Unit cost]]="",[@Quantity]=""),"",'
                  '[@[Unit cost]]*[@Quantity]))'
          },
          dropdowns={"Course": "Courses", "Status": "MenuStatus"},
          nonnegative=["Unit cost", "Quantity"])
    ws["Food and Drinks"].write("J8", "All menu items", f["section"])
    ws["Food and Drinks"].write_formula(
        "K8", "=SUM(Menu[Line total])", f["calc"])
    ws["Food and Drinks"].write("J9", "Selected only", f["section"])
    ws["Food and Drinks"].write_formula(
        "K9", '=SUMIF(Menu[Status],"Selected",Menu[Line total])', f["calc"])
    ws["Food and Drinks"].merge_range(
        "J12:L16",
        "Menu costing is an estimate, not an additional automatic budget "
        "charge. Record the agreed catering amount once in Wedding Budget "
        "to avoid double-counting.",
        f["wrap"],
    )

    table("Photoshoot", "Shots",
          ["Shot", "Category", "Location", "People", "Priority", "Status"],
          sample([
              ["Families together", "Family", "Garden", "Both families",
               "Must have", "Planned"],
              ["First look", "Couple", "Courtyard", f"{one} and {two}",
               "Must have", "Planned"],
              ["Sunset portraits", "Couple", "Terrace", f"{one} and {two}",
               "Nice to have", "Planned"],
          ]),
          dropdowns={"Category": "ShotCategory", "Priority": "Priority",
                     "Status": "ShotStatus"})

    table("Music", "MusicList",
          ["Song", "Artist", "Moment", "Notes"],
          sample([
              ["Our favourite song", "Chosen artist", "First dance",
               "Replace with the actual song and artist"],
              ["Ceremony entrance", "Chosen artist", "Walking in", ""],
          ]),
          dropdowns={"Moment": "MusicMoments"})

    table("Gifts and Thank You", "Gifts",
          ["From", "Gift", "Received", "Thank you", "Sent date", "Notes"],
          sample([
              ["Morgan household", "Dinnerware", wedding - timedelta(days=10),
               "Written", None, ""],
              ["Lee household", "Travel contribution",
               wedding - timedelta(days=5), "To write", None, ""],
          ]),
          styles={"Received": "date", "Sent date": "date"},
          dropdowns={"Thank you": "Thanks"})
    ws["Gifts and Thank You"].write("I8", "Still to send", f["section"])
    ws["Gifts and Thank You"].write_formula(
        "J8", '=COUNTIFS(Gifts[From],"<>",Gifts[Thank you],"<>Sent")',
        f["calc"],
    )

    # Dynamic seating: guest assignments have a single source of truth.
    seating = ws["Seating Plan"]
    seating.write("B6", "Selected table", f["text"])
    seating.write_string("C6", "1", f["input"])
    seating.data_validation("C6", {"validate": "list", "source": "=TableNames"})
    seating.write("B8", "Confirmed guests at selected table", f["section"])
    seating.write("E8", "Confirmed but not seated", f["section"])
    seating.write_dynamic_array_formula(
        "B10",
        '=FILTER(Guests[Full name]&" / "&Guests[Meal],'
        '(Guests[Full name]<>"")*(Guests[RSVP]="Attending")*'
        '(Guests[Table]&""=$C$6&""),"No confirmed guests")',
        f["calc_text"],
    )
    seating.write_dynamic_array_formula(
        "E10",
        '=FILTER(Guests[Full name],'
        '(Guests[Full name]<>"")*(Guests[RSVP]="Attending")*'
        '(Guests[Table]=""),"All confirmed guests seated")',
        f["calc_text"],
    )
    seating.add_table("H8:K28", {
        "name": "SeatingTables", "style": "Table Style Medium 4",
        "data": [[str(i), 8, None, None] for i in range(1, 21)],
        "columns": [
            {"header": "Table", "format": f["calc_text"]},
            {"header": "Capacity", "format": f["integer"]},
            {"header": "Attending", "format": f["calc_int"],
             "formula": '=COUNTIFS(Guests[Table],[@Table],'
                        'Guests[RSVP],"Attending",Guests[Full name],"<>")'},
            {"header": "Spaces left", "format": f["calc_int"],
             "formula": "=[@Capacity]-[@Attending]"},
        ],
    })
    seating.data_validation("I9:I28", {
        "validate": "integer", "criteria": ">=", "value": 0,
    })
    seating.conditional_format("K9:K28", {
        "type": "cell", "criteria": "<", "value": 0,
        "format": f["warning"],
    })

    # Macro-free Theme Board:
    # dropdown selection -> hex lookup -> conditional-format swatch.
    theme = ws["Theme Board"]

    theme.merge_range(
        "B6:I7",
        "Choose each palette colour from the dropdown above its swatch. "
        "The swatch and hex code update automatically. "
        "Insert inspiration images using Insert > Pictures > This Device. "
        "Changing a fill manually will not update its hex code.",
        f["wrap"],
    )

    default_names = ["Sage", "Blush", "Ivory"]
    palette_lookup = dict(THEME_PALETTE)

    swatch_base = wb.add_format({
        "align": "center",
        "valign": "vcenter",
        "border": 1,
    })

    # Reuse these formats across all three swatches.
    swatch_colour_formats = {
        hex_code: wb.add_format({"bg_color": hex_code})
        for _, hex_code in THEME_PALETTE
    }

    for i, default_name in enumerate(default_names):
        col = 1 + i * 3

        # Dropdowns sit above the swatches: B9, E9, H9.
        selector_cell = xlsxwriter.utility.xl_rowcol_to_cell(8, col)

        # Hex values sit underneath: B14, E14, H14.
        hex_cell = xlsxwriter.utility.xl_rowcol_to_cell(
            13, col, row_abs=True, col_abs=True
        )

        theme.write_string(
            8, col, default_name, f["input"]
        )

        theme.data_validation(8, col, 8, col, {
            "validate": "list",
            "source": "=ThemeColourNames",
            "input_title": f"Palette colour {i + 1}",
            "input_message": "Choose a colour from the dropdown.",
            "error_type": "stop",
            "error_title": "Choose a listed colour",
            "error_message": "Please select a colour from the dropdown.",
        })

        # Blank merged panel: its fill is controlled by the rules below.
        theme.merge_range(
            9, col, 12, col + 1,
            "",
            swatch_base,
        )

        # Read the selected colour's hex code.
        hex_formula = (
            f'=IFERROR(INDEX(ThemeColourHex,'
            f'MATCH({selector_cell},ThemeColourNames,0)),"")'
        )

        theme.write_formula(
            13, col,
            hex_formula,
            f["calc_text"],
            palette_lookup[default_name],
        )

        # Each predefined colour needs a conditional-formatting rule.
        # Excel cannot use a hex string directly as a fill colour.
        for _, hex_code in THEME_PALETTE:
            theme.conditional_format(
                9, col, 12, col + 1,
                {
                    "type": "formula",
                    "criteria": f'={hex_cell}="{hex_code}"',
                    "format": swatch_colour_formats[hex_code],
                },
            )

        # Keep the existing Setup named cells synchronized.
        # Their original field positions are D16, D17 and D18.
        setup.write_formula(
            15 + i, 3,
            f"='Theme Board'!{hex_cell}",
            f["calc_text"],
            palette_lookup[default_name],
        )

    for row, label in [(16, "Overall feeling"), (20, "Decor"),
                       (24, "Flowers"), (28, "Stationery")]:
        theme.write(row, 1, label, f["section"])
        theme.merge_range(row + 1, 1, row + 2, 8,
                          "Natural textures and warm light" if demo else "",
                          f["input"])

    gallery = ws["Photo Gallery"]
    gallery.merge_range(
        "B6:I7",
        "Insert > Pictures > This Device to embed a photo over each panel. "
        "Delete the panel text first if needed. Add titles and captions below. "
        "These panels are not upload buttons. Large images increase file size.",
        f["wrap"],
    )
    for row in [9, 24]:
        for col in [1, 6]:
            gallery.merge_range(row, col, row + 9, col + 3,
                                "Insert a local photo here", f["photo"])
            gallery.merge_range(row + 10, col, row + 10, col + 3,
                                "Photo title" if demo else "", f["input"])
            gallery.merge_range(row + 11, col, row + 12, col + 3,
                                "Caption" if demo else "", f["input"])
    gallery.print_area("B1:J38")

    # Calendar: six weeks; counts and balances use underlying tables.
    calendar = ws["Calendar"]
    calendar.write("B6", "Month — enter any date", f["text"])
    calendar.write_datetime("D6", wedding.replace(day=1), f["date"])
    calendar.data_validation("D6", {
        "validate": "date", "criteria": "between",
        "minimum": D(2020, 1, 1), "maximum": D(2100, 12, 31),
    })
    for col, day in enumerate(
        ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
         "Saturday", "Sunday"], start=1
    ):
        calendar.write(7, col, day, f["section"])

    day_format = wb.add_format({
        "num_format": "d mmm", "bg_color": blush, "bold": True,
    })
    for week in range(6):
        row = 8 + week * 5
        for day in range(7):
            col = 1 + day
            offset = week * 7 + day
            calendar.write_formula(
                row, col,
                '=DATE(YEAR($D$6),MONTH($D$6),1)'
                '-WEEKDAY(DATE(YEAR($D$6),MONTH($D$6),1),2)'
                f'+1+{offset}',
                day_format,
            )
            cell = xlsxwriter.utility.xl_rowcol_to_cell(row, col)
            formulas = [
                f'="Tasks: "&COUNTIFS(Tasks[Due date],{cell},'
                'Tasks[Status],"<>Done",Tasks[Task],"<>")',
                f'="Invites: "&COUNTIFS(Invites[Send by],{cell},'
                'Invites[Status],"<>Sent",Invites[Household],"<>")',
                f'="Visits: "&COUNTIFS(Venues[Viewing date],{cell},'
                'Venues[Venue],"<>")',
                f'="Due: "&TEXT(SUMIFS(Budget[Outstanding],'
                f'Budget[Payment due],{cell})+'
                f'SUMIFS(Honeymoon[Outstanding],'
                f'Honeymoon[Payment due],{cell}),"#,##0.00")',
            ]
            for line, formula in enumerate(formulas, start=1):
                calendar.write_formula(row + line, col, formula, f["text"])
    calendar.print_area("B1:H38")
    calendar.write(
        "J8", "Due includes wedding + honeymoon balances.", f["text"])

    # Live dashboard.
    dashboard = ws["Dashboard"]
    metrics = [
        ("Wedding budget", '=IF(WeddingBudget="","",WeddingBudget)'),
        ("Planned wedding spend", "=SUM(Budget[Planned cost])"),
        ("Paid so far", "=SUM(Budget[Paid to date])"),
        ("Outstanding", "=SUM(Budget[Outstanding])"),
        ("Budget remaining",
         '=IF(WeddingBudget="","",WeddingBudget-SUM(Budget[Planned cost]))'),
        ("Days to wedding", '=IF(WeddingDate="","",MAX(0,WeddingDate-TODAY()))'),
        ("Guests listed", '=COUNTIF(Guests[Full name],"<>")'),
        ("Attending", '=COUNTIFS(Guests[Full name],"<>",Guests[RSVP],"Attending")'),
        ("Awaiting response",
         '=COUNTIFS(Guests[Full name],"<>",Guests[RSVP],"Awaiting reply")'
         '+COUNTIFS(Guests[Full name],"<>",Guests[RSVP],"")'),
        ("Declined", '=COUNTIFS(Guests[Full name],"<>",Guests[RSVP],"Declined")'),
        ("Tasks complete", '=COUNTIFS(Tasks[Task],"<>",Tasks[Status],"Done")'),
        ("Tasks total", '=COUNTIF(Tasks[Task],"<>")'),
        ("Open tasks missing dates",
         '=COUNTIFS(Tasks[Task],"<>",Tasks[Status],"<>Done",Tasks[Due date],"")'),
        ("Expenses missing cost",
         '=COUNTIFS(Budget[Expense],"<>",Budget[Planned cost],"")'),
        ("Assignments on non-attendees",
         '=COUNTIFS(Guests[Full name],"<>",Guests[Table],"<>",'
         'Guests[RSVP],"<>Attending")'),
    ]
    for row, (label, formula) in enumerate(metrics, start=7):
        dashboard.write(row, 1, label, f["text"])
        number_format = f["calc"] if row < 12 else f["calc_int"]
        dashboard.write_formula(row, 3, formula, number_format)

    dashboard.write("B24", "Planning completion", f["section"])
    dashboard.write_formula(
        "D24", '=IFERROR(COUNTIFS(Tasks[Task],"<>",Tasks[Status],"Done")/'
               'COUNTIF(Tasks[Task],"<>"),0)', f["percent"])

    # Supporting category totals, below the chart.
    dashboard.write("J27", "Category", f["section"])
    dashboard.write("K27", "Planned", f["section"])

    for row, category in enumerate(CATEGORIES, start=27):
        dashboard.write(row, 9, category, f["text"])
        dashboard.write_formula(
            row, 10,
            f'=SUMIF(Budget[Category],J{row+1},Budget[Planned cost])',
            f["calc"],
        )

    chart = wb.add_chart({"type": "doughnut"})
    chart.add_series({
        "name": "Planned wedding spend",
        "categories": "='Dashboard'!$J$28:$J$37",
        "values": "='Dashboard'!$K$28:$K$37",
        "points": [
            {"fill": {"color": colour}}
            for colour in [
                "#81917A", "#A6B49F", "#C5CEBD", "#D7B8AA",
                "#E9D1C8", "#AFBCAC", "#B5A697", "#DED9CD",
                "#9AAB99", "#C8BEAE",
            ]
        ],
    })
    chart.set_title({"name": "Planned spend by category"})
    chart.set_legend({"position": "bottom"})
    chart.set_chartarea({
        "border": {"none": True},
        "fill": {"color": "#FAF8F2"},
    })
    chart.set_size({"width": 540, "height": 310})
    dashboard.insert_chart("F8", chart)

    dashboard.set_zoom(100)

    dashboard.write("B28", "Open tasks, ordered by due date", f["section"])
    dashboard.write_dynamic_array_formula(
        "B30",
        '=SORT(FILTER(Tasks[[Task]:[Due date]],'
        '(Tasks[Task]<>"")*(Tasks[Status]<>"Done")*'
        '(Tasks[Due date]<>""),"No dated open tasks"),6,1)',
        f["calc_text"],
    )
    dashboard.set_column("G:G", 24)

    # Apply date formatting only to dates in the upcoming-task list.
    dashboard.conditional_format("F30:G5000", {
        "type": "formula",
        "criteria": "=ISNUMBER(F30)",
        "format": f["calc_date"],
    })

    dashboard.print_area("B1:L40")

    # Start Here, navigation, boundaries and honest limitations.
    start.merge_range(
        "B6:J8",
        "ASSESSMENT ONLY — Adaptation of Johan Botha's wedding-planner "
        "reference pack. Supplied materials remain subject to the assessment "
        "licence. This workbook is not authorised for independent resale.",
        f["wrap"],
    )
    instructions = [
        "1. Keep the clean master unchanged. Use File > Save As for each wedding.",
        "2. Complete Setup: names, wedding date, venue, budgets and currency.",
        "3. Amber cells are inputs; pale green cells contain calculated results.",
        "4. Use dropdowns and enter genuine Excel dates/times, not date-like text.",
        "5. Add records with Tab from the last table cell; do not type below a gap.",
        "6. Guest List is the source of truth for RSVPs, meals and table assignments.",
        "7. Budget and menu are separate estimates: do not duplicate catering costs.",
        "8. Tasks use days before the wedding unless a Fixed date overrides them.",
        "9. Sort itinerary by Time before printing. Review Print Preview first.",
        "10. To start again safely, create another copy of the clean master.",
        "11. No external links, account, macros or Python are needed to use this file.",
        "12. Personal information remains in this file; share and store it carefully.",
    ]
    for row, text in enumerate(instructions, start=10):
        start.merge_range(row, 1, row, 11, text, f["text"])

    start.merge_range(
        "B25:L28",
        "Design limits: 20 seating tables; 4 gallery panels; dropdown/date/"
        "numeric validation pre-applied through worksheet row 5000. Native "
        "Excel Tables and their formulas can grow, but extend validation when "
        "exceeding row 5000. Do not place content below dynamic seating/task "
        "lists: they need clear space to expand. Formula cells are not locked "
        "so table expansion remains straightforward. Standard Excel copy/paste "
        "can bypass validation. Review print areas after adding records.",
        f["wrap"],
    )
    start.merge_range(
        "B30:L32",
        "Currency: all numbers use the Setup currency label; no exchange-rate "
        "conversion. Payment model: one cumulative paid amount and one "
        "remaining-balance due date per expense, not an instalment ledger. "
        "Calendar visits include historical viewings. Seating shows confirmed "
        "attendees only. Images are inserted manually. Theme colours are "
        "selected from predefined dropdowns on Theme Board.",
        f["wrap"],
    )
    start.write("B35", "Planning pages / reference coverage", f["section"])
    for i, name in enumerate(SHEETS, start=1):
        start.write_url(
            35 + i, 1, f"internal:'{name}'!B1",
            f["link"], f"{i:02}  {name}",
        )
    start.merge_range(
        "B59:L61",
        "Tested by the candidate in Microsoft Excel for Microsoft 365 on "
        "Windows, 64-bit, Version 2610, Build 16.0.20527.20196. "
        "See validation.md for test results and limitations. Other versions "
        "and platforms have not been verified. Excel settings control decimal "
        "and thousands separators. Python is needed only to generate the "
        "files, not to use them.",
        f["wrap"],
    )
    start.print_area("B1:L62")
    start.activate()
    start.set_first_sheet()
    wb.close()
    print(f"Created: {OUTPUT / client['file']}")


if __name__ == "__main__":
    for client in CLIENTS:
        build(client)
    print("\nGeneration complete. Open each file in Excel to recalculate.")