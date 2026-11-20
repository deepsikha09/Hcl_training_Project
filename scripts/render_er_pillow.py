import os
from PIL import Image, ImageDraw, ImageFont

def render_er_diagram_pillow(output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # 2400 x 1700 canvas
    width = 2400
    height = 1750
    img = Image.new('RGB', (width, height), color='#FFFFFF')
    draw = ImageDraw.Draw(img)
    
    # Fonts - try standard Windows TrueType fonts
    try:
        font_title = ImageFont.truetype("arialbd.ttf", 36)
        font_sub = ImageFont.truetype("ariali.ttf", 20)
        font_tbl_hdr = ImageFont.truetype("arialbd.ttf", 20)
        font_col = ImageFont.truetype("arial.ttf", 17)
        font_col_bold = ImageFont.truetype("arialbd.ttf", 17)
        font_type = ImageFont.truetype("consola.ttf", 15)
        font_badge = ImageFont.truetype("arialbd.ttf", 14)
        font_link = ImageFont.truetype("arialbd.ttf", 17)
    except:
        font_title = font_sub = font_tbl_hdr = font_col = font_col_bold = font_type = font_badge = font_link = ImageFont.load_default()

    # Title Banner
    draw.text((width // 2, 45), "APEXCARE HEALTH NETWORK — ENTITY RELATIONSHIP (ER) DIAGRAM", 
              fill="#1F497D", font=font_title, anchor="mm")
    draw.text((width // 2, 85), "Physical Relational Data Model (Normalized Star/Snowflake Schema) | Step 09 Specification", 
              fill="#64748B", font=font_sub, anchor="mm")
    draw.line([(100, 115), (width - 100, 115)], fill="#CBD5E1", width=2)

    # Color definitions
    c_navy_hdr = "#1F497D"
    c_fact_hdr = "#1B365D"
    c_card_bg = "#F8FAFC"
    c_card_border = "#94A3B8"
    c_stripe = "#F1F5F9"
    c_pk_bg = "#FFEDD5"
    c_pk_text = "#C2410C"
    c_fk_bg = "#D1FAE5"
    c_fk_text = "#047857"
    c_uk_bg = "#EDE9FE"
    c_uk_text = "#6366F1"

    def draw_table_card(x, y, w, h, title, columns, is_fact=False):
        # Card shadow
        draw.rectangle([x+4, y+4, x+w+4, y+h+4], fill="#E2E8F0")
        # Card Base
        draw.rectangle([x, y, x+w, y+h], fill=c_card_bg, outline=c_card_border, width=2)
        
        # Header
        hdr_h = 48
        hdr_color = c_fact_hdr if is_fact else c_navy_hdr
        draw.rectangle([x, y, x+w, y+hdr_h], fill=hdr_color)
        
        badge_type = "[FACT TABLE]" if is_fact else "[MASTER DIMENSION]"
        draw.text((x + w//2, y + hdr_h//2), f"{title.upper()}  {badge_type}", 
                  fill="#FFFFFF", font=font_tbl_hdr, anchor="mm")
        
        # Rows
        row_y = y + hdr_h
        row_h = 32
        for i, (col_name, col_type, col_key) in enumerate(columns):
            r_top = row_y + (i * row_h)
            r_bot = r_top + row_h
            if i % 2 == 1:
                draw.rectangle([x+1, r_top, x+w-1, r_bot], fill=c_stripe)
            
            # Key Badge
            if col_key == "PK":
                draw.rectangle([x+10, r_top+5, x+48, r_bot-5], fill=c_pk_bg, outline=c_pk_text, width=1)
                draw.text((x+29, r_top+row_h//2), "PK", fill=c_pk_text, font=font_badge, anchor="mm")
            elif col_key == "FK":
                draw.rectangle([x+10, r_top+5, x+48, r_bot-5], fill=c_fk_bg, outline=c_fk_text, width=1)
                draw.text((x+29, r_top+row_h//2), "FK", fill=c_fk_text, font=font_badge, anchor="mm")
            elif col_key == "UK":
                draw.rectangle([x+10, r_top+5, x+48, r_bot-5], fill=c_uk_bg, outline=c_uk_text, width=1)
                draw.text((x+29, r_top+row_h//2), "UK", fill=c_uk_text, font=font_badge, anchor="mm")
            else:
                draw.ellipse([x+26, r_top+13, x+32, r_top+19], fill="#94A3B8")

            # Column Name
            f_name = font_col_bold if col_key else font_col
            draw.text((x + 60, r_top + row_h//2), col_name, fill="#1E293B", font=f_name, anchor="lm")
            
            # Column Type
            draw.text((x + w - 15, r_top + row_h//2), col_type, fill="#64748B", font=font_type, anchor="rm")
            
            # Sub-border
            draw.line([(x, r_bot), (x+w, r_bot)], fill="#E2E8F0", width=1)

    # Columns definitions
    cols_clinics = [
        ("clinic_id", "INTEGER", "PK"),
        ("clinic_code", "TEXT", "UK"),
        ("clinic_name", "TEXT", ""),
        ("location_sector", "TEXT", ""),
        ("city", "TEXT", ""),
        ("contact_number", "TEXT", ""),
        ("daily_slot_capacity", "INTEGER", "")
    ]
    
    cols_specialties = [
        ("specialty_id", "INTEGER", "PK"),
        ("specialty_code", "TEXT", "UK"),
        ("specialty_name", "TEXT", ""),
        ("standard_consultation_fee", "REAL", ""),
        ("standard_slot_duration_mins", "INTEGER", ""),
        ("department_head", "TEXT", "")
    ]
    
    cols_doctors = [
        ("doctor_id", "INTEGER", "PK"),
        ("doctor_code", "TEXT", "UK"),
        ("doctor_name", "TEXT", ""),
        ("specialty_id", "INTEGER", "FK"),
        ("clinic_id", "INTEGER", "FK"),
        ("qualification", "TEXT", ""),
        ("hourly_idle_rate", "REAL", "")
    ]
    
    cols_patients = [
        ("patient_id", "INTEGER", "PK"),
        ("patient_mrn", "TEXT", "UK"),
        ("age", "INTEGER", ""),
        ("age_group", "TEXT", ""),
        ("gender", "TEXT", ""),
        ("neighborhood", "TEXT", ""),
        ("transit_distance_km", "REAL", ""),
        ("hypertension_flag", "INTEGER", ""),
        ("diabetes_flag", "INTEGER", ""),
        ("chronic_comorbidity_score", "INTEGER", ""),
        ("corporate_partner_flag", "INTEGER", ""),
        ("corporate_shift_type", "TEXT", ""),
        ("registration_date", "TEXT", "")
    ]
    
    cols_appointments = [
        ("appointment_id", "INTEGER", "PK"),
        ("booking_reference", "TEXT", "UK"),
        ("patient_id", "INTEGER", "FK"),
        ("doctor_id", "INTEGER", "FK"),
        ("specialty_id", "INTEGER", "FK"),
        ("clinic_id", "INTEGER", "FK"),
        ("booking_date", "TEXT", ""),
        ("appointment_date", "TEXT", ""),
        ("appointment_time", "TEXT", ""),
        ("lead_time_days", "INTEGER", ""),
        ("lead_time_bucket", "TEXT", ""),
        ("appointment_day_of_week", "TEXT", ""),
        ("appointment_time_slot", "TEXT", ""),
        ("sms_reminder_sent", "INTEGER", ""),
        ("reminder_count", "INTEGER", ""),
        ("prior_appointments_count", "INTEGER", ""),
        ("prior_no_shows_count", "INTEGER", ""),
        ("repeat_offender_tier", "TEXT", ""),
        ("actual_consultation_fee", "REAL", ""),
        ("scheduled_duration_mins", "INTEGER", ""),
        ("status", "TEXT", ""),
        ("is_no_show", "INTEGER", ""),
        ("risk_score", "REAL", ""),
        ("risk_tier", "TEXT", ""),
        ("capacity_loss_hours", "REAL", ""),
        ("financial_loss_inr", "REAL", "")
    ]

    # Layout Coordinates
    # dim_clinics: Top Left
    draw_table_card(80, 160, 480, 48 + 7*32, "dim_clinics", cols_clinics)
    
    # dim_specialties: Middle Left
    draw_table_card(80, 520, 480, 48 + 6*32, "dim_specialties", cols_specialties)
    
    # dim_doctors: Bottom Left
    draw_table_card(80, 840, 480, 48 + 7*32, "dim_doctors", cols_doctors)
    
    # fact_appointments: Center
    draw_table_card(780, 160, 780, 48 + 26*32, "fact_appointments", cols_appointments, is_fact=True)
    
    # dim_patients: Top Right
    draw_table_card(1780, 160, 540, 48 + 13*32, "dim_patients", cols_patients)

    # Relationships Connectors & Badges
    def draw_link(p_start, p_end, label):
        # Line
        draw.line([p_start, p_end], fill="#334155", width=3)
        # Arrowhead
        x2, y2 = p_end
        if p_start[0] < p_end[0]: # pointing right
            draw.polygon([(x2, y2), (x2-12, y2-7), (x2-12, y2+7)], fill="#334155")
        elif p_start[0] > p_end[0]: # pointing left
            draw.polygon([(x2, y2), (x2+12, y2-7), (x2+12, y2+7)], fill="#334155")
        elif p_start[1] < p_end[1]: # pointing down
            draw.polygon([(x2, y2), (x2-7, y2-12), (x2+7, y2-12)], fill="#334155")
            
        # Label Box
        mid_x = (p_start[0] + p_end[0]) // 2
        mid_y = (p_start[1] + p_end[1]) // 2
        
        box_w = 170
        box_h = 32
        draw.rectangle([mid_x - box_w//2, mid_y - box_h//2, mid_x + box_w//2, mid_y + box_h//2], 
                       fill="#FFFFFF", outline="#475569", width=2)
        draw.text((mid_x, mid_y), label, fill="#0F172A", font=font_link, anchor="mm")

    # 1. dim_clinics -> fact_appointments
    draw_link((560, 290), (780, 290), "hosts (1:M)")
    
    # 2. dim_specialties -> fact_appointments
    draw_link((560, 630), (780, 630), "categorizes (1:M)")
    
    # 3. dim_doctors -> fact_appointments
    draw_link((560, 950), (780, 950), "conducts (1:M)")
    
    # 4. dim_patients -> fact_appointments
    draw_link((1780, 360), (1560, 360), "books (1:M)")
    
    # 5. dim_specialties -> dim_doctors (downward)
    draw_link((320, 760), (320, 840), "employs (1:M)")

    # 6. dim_clinics -> dim_doctors (left curve)
    draw.line([(80, 300), (30, 300), (30, 960), (80, 960)], fill="#334155", width=3)
    draw.polygon([(80, 960), (68, 953), (68, 967)], fill="#334155")
    draw.rectangle([10, 615, 120, 647], fill="#FFFFFF", outline="#475569", width=2)
    draw.text((65, 631), "stations (1:M)", fill="#0F172A", font=font_badge, anchor="mm")

    # Notation Legend Card
    leg_x = 1780
    leg_y = 680
    leg_w = 540
    leg_h = 380
    draw.rectangle([leg_x+4, leg_y+4, leg_x+leg_w+4, leg_y+leg_h+4], fill="#E2E8F0")
    draw.rectangle([leg_x, leg_y, leg_x+leg_w, leg_y+leg_h], fill="#FFFFFF", outline="#1F497D", width=2)
    draw.rectangle([leg_x, leg_y, leg_x+leg_w, leg_y+42], fill="#1F497D")
    draw.text((leg_x + leg_w//2, leg_y + 21), "ER DIAGRAM NOTATION KEY", fill="#FFFFFF", font=font_tbl_hdr, anchor="mm")
    
    items = [
        ("PK Badge (Orange)", "Primary Key (Uniquely identifies table record)"),
        ("FK Badge (Green)", "Foreign Key Reference (Enforces relational integrity)"),
        ("UK Badge (Purple)", "Unique Business Key (Alternate unique constraint)"),
        ("1 : M Connectors", "One-to-Many Crow's Foot Relational Cardinality"),
        ("ON DELETE RESTRICT", "Prevents orphan fact records upon master delete"),
        ("CHECK Constraints", "Enforces domain validity (e.g., status in Show/No-Show)"),
        ("vw_appointment_analytics", "Denormalized 36-column view bridging to Pandas")
    ]
    for idx, (title, desc) in enumerate(items):
        it_y = leg_y + 65 + (idx * 42)
        draw.text((leg_x + 20, it_y), f"•  {title}:", fill="#1F497D", font=font_col_bold, anchor="lm")
        draw.text((leg_x + 230, it_y), desc, fill="#334155", font=font_col, anchor="lm")

    img.save(output_path, "PNG", quality=95)
    print(f"Pillow-rendered ER Diagram successfully saved at: {output_path}")

if __name__ == "__main__":
    out_file = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\09_ER_Diagram.png"
    render_er_diagram_pillow(out_file)
