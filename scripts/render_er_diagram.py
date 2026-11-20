import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches

def generate_er_image(output_png):
    os.makedirs(os.path.dirname(output_png), exist_ok=True)
    
    # High-resolution canvas: 16 x 11 inches at 300 dpi
    fig, ax = plt.subplots(figsize=(16, 11), dpi=300)
    ax.set_xlim(0, 16)
    ax.set_ylim(0, 11)
    ax.axis('off')
    
    # Colors
    c_header = '#1F497D'      # Deep Navy
    c_fact_header = '#1B365D' # Darker Navy for Fact Table
    c_bg = '#F8FAFC'          # Light Gray/Blue background for tables
    c_border = '#CBD5E1'      # Slate border
    c_pk = '#C2410C'          # Terracotta Orange for PK
    c_fk = '#047857'          # Forest Green for FK
    c_line = '#475569'        # Dark slate for relationship lines
    
    # Title Block
    ax.text(8.0, 10.6, "APEXCARE HEALTH NETWORK — ENTITY RELATIONSHIP (ER) DIAGRAM", 
            fontsize=15, fontweight='bold', color='#1F497D', ha='center', va='center', fontfamily='sans-serif')
    ax.text(8.0, 10.25, "Physical Relational Data Model (Normalized Star/Snowflake Schema) | Step 09 Specification", 
            fontsize=10.5, fontstyle='italic', color='#64748B', ha='center', va='center', fontfamily='sans-serif')

    def draw_entity_table(x, y, w, h, title, columns, is_fact=False):
        header_h = 0.55
        hdr_color = c_fact_header if is_fact else c_header
        
        # Shadow
        shadow = patches.FancyBboxPatch((x+0.04, y-0.04), w, h, boxstyle="round,pad=0.04", 
                                        fc='#E2E8F0', ec='none', zorder=1)
        ax.add_patch(shadow)
        
        # Base Card
        card = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.04", 
                                      fc=c_bg, ec=c_border, lw=1.2, zorder=2)
        ax.add_patch(card)
        
        # Header Box
        header_box = patches.FancyBboxPatch((x, y + h - header_h), w, header_h, 
                                            boxstyle="round,pad=0.04", fc=hdr_color, ec='none', zorder=3)
        ax.add_patch(header_box)
        
        # Title text
        badge = " [FACT TABLE]" if is_fact else " [DIMENSION]"
        ax.text(x + w/2.0, y + h - header_h/2.0, title.upper() + badge, 
                fontsize=9.5, fontweight='bold', color='white', ha='center', va='center', zorder=4)
        
        # Render columns
        curr_y = y + h - header_h - 0.22
        line_spacing = 0.23
        for col_name, col_type, col_key in columns:
            # Key indicator
            if col_key == "PK":
                ax.text(x + 0.15, curr_y, "PK", fontsize=7.5, fontweight='bold', color=c_pk, va='center', zorder=4)
            elif col_key == "FK":
                ax.text(x + 0.15, curr_y, "FK", fontsize=7.5, fontweight='bold', color=c_fk, va='center', zorder=4)
            elif col_key == "UK":
                ax.text(x + 0.15, curr_y, "UK", fontsize=7.5, fontweight='bold', color='#6366F1', va='center', zorder=4)
            else:
                ax.text(x + 0.15, curr_y, "•", fontsize=8.0, color='#94A3B8', va='center', zorder=4)
                
            # Column Name
            ax.text(x + 0.55, curr_y, col_name, fontsize=8.2, fontweight='bold' if col_key else 'normal', 
                    color='#1E293B', va='center', zorder=4)
            
            # Type
            ax.text(x + w - 0.15, curr_y, col_type, fontsize=7.5, color='#64748B', ha='right', va='center', zorder=4)
            
            curr_y -= line_spacing

    # Table Definitions
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
    
    # Position Tables
    # dim_clinics (Top Left)
    draw_entity_table(0.6, 7.3, 3.2, 2.5, "dim_clinics", cols_clinics)
    
    # dim_specialties (Bottom Left)
    draw_entity_table(0.6, 3.8, 3.2, 2.3, "dim_specialties", cols_specialties)
    
    # dim_doctors (Middle Left)
    draw_entity_table(0.6, 0.4, 3.2, 2.5, "dim_doctors", cols_doctors)
    
    # dim_patients (Top Right)
    draw_entity_table(11.8, 5.0, 3.6, 4.0, "dim_patients", cols_patients)
    
    # fact_appointments (Center)
    draw_entity_table(4.8, 0.5, 6.0, 9.3, "fact_appointments", cols_appointments, is_fact=True)
    
    # Relationship Connectors & Annotations
    def draw_connector(p1, p2, label, label_pos, text_offset=(0, 0)):
        # Line
        ax.annotate('', xy=p2, xytext=p1,
                    arrowprops=dict(arrowstyle="->", color=c_line, lw=1.6, 
                                    shrinkA=4, shrinkB=4, connectionstyle="arc3,rad=0.05"),
                    zorder=5)
        # Verb Label
        ax.text(label_pos[0] + text_offset[0], label_pos[1] + text_offset[1], 
                f"  {label} (1:M)  ", fontsize=7.8, fontweight='bold', color='#0F172A',
                bbox=dict(boxstyle="round,pad=0.25", fc='#FFFFFF', ec='#94A3B8', lw=0.9),
                ha='center', va='center', zorder=6)

    # 1. dim_clinics -> fact_appointments
    draw_connector((3.8, 8.5), (4.8, 8.5), "hosts", (4.3, 8.5))
    
    # 2. dim_specialties -> fact_appointments
    draw_connector((3.8, 5.0), (4.8, 5.0), "categorizes", (4.3, 5.0))
    
    # 3. dim_doctors -> fact_appointments
    draw_connector((3.8, 1.8), (4.8, 1.8), "conducts", (4.3, 1.8))
    
    # 4. dim_patients -> fact_appointments
    draw_connector((11.8, 6.8), (10.8, 6.8), "books", (11.3, 6.8))
    
    # 5. dim_specialties -> dim_doctors
    draw_connector((2.2, 3.8), (2.2, 2.9), "employs", (2.2, 3.35), text_offset=(0.45, 0))

    # 6. dim_clinics -> dim_doctors (Curved connection down)
    ax.annotate('', xy=(0.6, 1.8), xytext=(0.6, 7.3),
                arrowprops=dict(arrowstyle="->", color=c_line, lw=1.3, linestyle='--',
                                connectionstyle="arc3,rad=0.3"),
                zorder=5)
    ax.text(0.12, 4.5, "stations\n(1:M)", fontsize=7.2, fontweight='bold', color='#0F172A',
            bbox=dict(boxstyle="round,pad=0.2", fc='#FFFFFF', ec='#94A3B8', lw=0.8),
            ha='center', va='center', zorder=6)

    # Legend / Cardinality Key
    leg_box = patches.FancyBboxPatch((11.8, 0.8), 3.6, 2.8, boxstyle="round,pad=0.04", 
                                     fc='#FFFFFF', ec='#94A3B8', lw=1.0, zorder=2)
    ax.add_patch(leg_box)
    ax.text(13.6, 3.25, "ER DIAGRAM NOTATION KEY", fontsize=8.5, fontweight='bold', color='#1F497D', ha='center', va='center')
    
    ax.text(12.0, 2.8, "PK", fontsize=7.5, fontweight='bold', color=c_pk)
    ax.text(12.5, 2.8, ": Primary Key Attribute", fontsize=7.8, color='#333333')
    
    ax.text(12.0, 2.4, "FK", fontsize=7.5, fontweight='bold', color=c_fk)
    ax.text(12.5, 2.4, ": Foreign Key Reference", fontsize=7.8, color='#333333')
    
    ax.text(12.0, 2.0, "UK", fontsize=7.5, fontweight='bold', color='#6366F1')
    ax.text(12.5, 2.0, ": Unique Business Key", fontsize=7.8, color='#333333')
    
    ax.text(12.0, 1.6, "1:M", fontsize=7.5, fontweight='bold', color='#0F172A')
    ax.text(12.5, 1.6, ": One-to-Many Relationship", fontsize=7.8, color='#333333')
    
    ax.text(12.0, 1.2, "-->", fontsize=7.5, fontweight='bold', color=c_line)
    ax.text(12.5, 1.2, ": Foreign Key Constraint", fontsize=7.8, color='#333333')

    plt.tight_layout()
    plt.savefig(output_png, bbox_inches='tight', dpi=300)
    plt.close()
    print(f"High-resolution ER diagram PNG successfully saved at: {output_png}")

if __name__ == "__main__":
    out_file = r"c:\Users\deeps\Downloads\Appointment No Show Analysis\docs\09_ER_Diagram.png"
    generate_er_image(out_file)
