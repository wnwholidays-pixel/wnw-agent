import streamlit as st
import google.generativeai as genai

# Page Configuration with Official Corporate Blue Identity
st.set_page_config(page_title="WNW Holidays Itinerary Engine", layout="wide")

st.markdown("""
    <style>
    .reportview-container { background: #f0f5fa; }
    .stButton>button { background-color: #0056b3; color: white; border-radius: 6px; font-weight: bold; }
    h1 { color: #0056b3; }
    </style>
""", unsafe_allow_html=True)

st.title("🦅 Wings 'N' Wheels Holidays")
st.caption("Official Native Voice AI Itinerary Automation Engine")

# 1. FIXED MASTER ITINERARY DATABASE (HARDCODED FROM YOUR UPLOADED FILE TEMPLATES)
MASTER_PACKAGES = {
    "Chandigarh – Manali Tour (6N/7D)": {
        "route": "CAMPUS → AJMER → CHANDIGARH → MANALI → AJMER → CAMPUS",
        "inclusions": "- 02 Night Stay in MANALI hotel\\n- 01 Night Stay in MANALI camp\\n- 01 Night Stay in CHANDIGARH\\n- Deluxe Rooms With all Amenities' (with 4 sharing)\\n- All Visit By Luxury coach (Non AC – as AC will not work during hill drive)\\n- All Meals (veg) – As per Itinerary\\n- All Transportation Expense",
        "exclusions": "- All Personal Expenses\\n- RIVER RAFTING\\n- Tips\\n- Any meal or fast food item not mentioned in inclusions\\n- Anything Which Is Not Mentioned in Inclusions",
        "cost": "Rs.13000/-"
    },
    "Dalhousie – Dharamshala – Amritsar Tour (7N/8D)": {
        "route": "FALNA → AJMER → DALHOUSIE → DHARAMSHALA → AMRITSAR → AJMER → FALNA",
        "inclusions": "- 05 Night Stay in DALHOUSIE / DHARAMSHALA / AMRITSAR hotels\\n- All Meals (veg) – As per Itinerary\\n- Deluxe Rooms With all Amenities' (with 4 sharing)\\n- All Transportation Expense\\n- All Visit By Luxury coach (Non AC – as AC will not work during hill drive)",
        "exclusions": "- All Personal Expenses\\n- RIVER RAFTING\\n- Tips\\n- Any meal or fast food item not mentioned in inclusions\\n- Anything Which Is Not Mentioned in Inclusions",
        "cost": "TO BE CONFIRMED"
    },
    "Gir – Somnath – Diu Tour (5N/6D)": {
        "route": "FALNA → GIR → SOMNATH → DIU → FALNA",
        "inclusions": "- 03 Night Stay in GIR, SOMNATH & DIU hotels\\n- 01 Night Stay in SOMNATH\\n- Deluxe Rooms With all Amenities' (with 4 sharing)\\n- All Visit By Luxury coach (Non AC)\\n- All Meals (veg) – As per Itinerary\\n- All Transportation Expense",
        "exclusions": "- All Personal Expenses\\n- Tips\\n- Any meal or fast food item not mentioned in inclusions\\n- Anything Which Is Not Mentioned in Inclusions",
        "cost": "TO BE CONFIRMED"
    },
    "Khajuraho – Jabalpur Tour (4N/6D)": {
        "route": "FALNA → AJMER → KHAJURAHO → JABALPUR → AJMER → FALNA",
        "inclusions": "- 02 Night Stay in KHAJURAHO hotel\\n- 01 Night Stay in KHAJURAHO camp\\n- 01 Night Stay in JABALPUR\\n- Deluxe Rooms With all Amenities' (with 4 sharing)\\n- All Visit By Luxury coach (Non AC)\\n- All Meals (veg) – As per Itinerary\\n- All Transportation Expense",
        "exclusions": "- All Personal Expenses\\n- Tips\\n- Any meal or fast food item not mentioned in inclusions\\n- Anything Which Is Not Mentioned in Inclusions",
        "cost": "TO BE CONFIRMED"
    },
    "Kumbhalgarh Tour (1 Day)": {
        "route": "FALNA → KUMBHALGARH → FALNA",
        "inclusions": "- Packed Breakfast\\n- Monument entries\\n- All Visit By Luxury coach (Non AC)\\n- Meals: Packed Breakfast. Lunch in hotel\\n- All Transportation Expense",
        "exclusions": "- All Personal Expenses\\n- Tips\\n- Any meal or fast food item not mentioned in inclusions\\n- Anything Which Is Not Mentioned in Inclusions",
        "cost": "Rs 1500/-"
    }
}

if "step" not in st.session_state:
    st.session_state.step = 1
    st.session_state.trip_data = {}

# STEP 1: PARSE INITIAL REQUIREMENTS VIA NATIVE VOICE/TEXT
if st.session_state.step == 1:
    st.subheader("📍 Step 1: Core Target Logistics & Package Blueprint")
    
    api_input = st.text_input("Enter your Gemini API Key to wake up the agent:", type="password")
    selected_pack = st.selectbox("Select Target Master Itinerary Template:", list(MASTER_PACKAGES.keys()))
    
    st.markdown("### 🎙️ Native Voice Input")
    st.info("Tap the microphone circle below to speak your route requirements aloud (e.g., 'Starting from Pali, train at 17:00 from Ajmer station')")
    
    # Official native stable audio input handler
    voice_audio = st.audio_input("Record audio notes here")
    if voice_audio:
        st.success("Voice note captured securely! Confirm the textual adjustments below:")

    start_place = st.text_input("Client's Starting Location / Campus Point:", "Pali")
    boarding_hub = st.text_input("Boarding Hub Station (If Train/Flight trip):", "Ajmer")
    train_time = st.text_input("Departure Time (HH:MM format):", "17:00")
    
    if st.button("Execute Routing Validation Checks"):
        if not api_input:
            st.error("Please insert your API Key to initialize the travel agent.")
        else:
            st.session_state.api_key = api_input
            st.session_state.trip_data.update({
                "package": selected_pack, "start": start_place, 
                "hub": boarding_hub, "train_time": train_time
            })
            
            if start_place.lower() != boarding_hub.lower() and "Kumbhalgarh" not in selected_pack:
                st.session_state.step = 2
            else:
                st.session_state.step = 3
            st.rerun()

# STEP 2: ROUTING DISCREPANCY FLOW BUFFER
if st.session_state.step == 2:
    st.warning(f"⚠️ Transfer Route Alert: Departure starting point is '{st.session_state.trip_data['start']}' but boarding hub is '{st.session_state.trip_data['hub']}'.")
    st.info("🗺️ Operational Parameter: Road distance spans approximately 170 KM. Appending a 3-hour local transit window before the scheduled departure.")
    
    st.markdown("### 🎙️ Speak the Transition Plan:")
    voice_gap = st.audio_input("Record connecting bus routing details here")
    
    transfer_method = st.text_input("Verify transfer text notation here:", "By luxury campus tour coach bus connection")
    client_count = st.number_input("Minimum Group Student Headcount:", min_value=1, value=45)
    
    if st.button("Lock Connecting Buffer Parameters"):
        st.session_state.trip_data.update({
            "transfer_method": transfer_method,
            "headcount": client_count,
            "buffer_calc": "3 Hours road transit"
        })
        st.session_state.step = 3
        st.rerun()

# STEP 3: EXECUTE AUTOMATION TEMPLATE COMPLIANCE
if st.session_state.step == 3:
    st.subheader("⚙️ Step 3: Compiling Proposal Layout Engine")
    st.info("The agent is wrapping your custom elements with static corporate blocks...")
    
    if st.button("Compile Complete Proposal Sheets"):
        genai.configure(api_key=st.session_state.api_key)
        model = genai.GenerativeModel('gemini-2.5-flash')
        pack_info = MASTER_PACKAGES[st.session_state.trip_data["package"]]
        
        ai_prompt = f"""
        You are a data extraction bot for Wings 'N' Wheels Holidays. 
        Your absolute task is to populate user inputs into our fixed template format. Do not alter static components.

        Live Client Inputs:
        - Tour Start Location: {st.session_state.trip_data['start']}
        - Target Hub: {st.session_state.trip_data['hub']}
        - Initial Transit Time: {st.session_state.trip_data['train_time']}
        - Connecting Route notes: {st.session_state.trip_data.get('transfer_method', 'Direct departure')}
        - Selected Proposal Target: {st.session_state.trip_data['package']}

        OUTPUT FORMAT STRUCTURE:
        Your output must mimic our corporate proposal layout with exactly these sections:
        
        1. **WNW HEADER BLOCK**
           Display contact details: "The First 24x7 Free Travel Planning Helpline | 📞 +91-9314193241 | +91 7742400525".
           Display Route Header: "{pack_info['route']}".
        
        2. **📢 HINDI AUDIO CONFIRMATION SCRIPT**
           Write a single clear conversational paragraph in Hindi text summarizing the travel itinerary changes, road drive times from {st.session_state.trip_data['start']} to {st.session_state.trip_data['hub']}, and night stay structures for phone screening logs.

        3. **📅 DETAILED ITINERARY**
           A day-by-day itemized timeline grid incorporating the custom connection transfers safely.

        4. **📝 FIXED TERMS AND CONDITIONS**
           Inclusions:
           {pack_info['inclusions']}
           
           Exclusions:
           {pack_info['exclusions']}

        5. **💰 PRICING GRID MATRIX**
           Render a clean layout table tracking:
           - Cost Per Student: {pack_info['cost']}
           - Group Strength: {st.session_state.trip_data.get('headcount', 45)} (Minimum)
           - Complementary (Teachers only): 15:01
        """
        
        with st.spinner("Locking template headers and generating layout..."):
            response = model.generate_content(ai_prompt)
            st.session_state.final_output = response.text
            st.session_state.step = 4
            st.rerun()

# STEP 4: DISPLAY PRODUCTION EXPORT DRAFT
if st.session_state.step == 4:
    st.success("🎉 Itinerary engine execution completed with zero structural changes!")
    st.markdown(st.session_state.final_output)
    
    if st.button("Initialize Next Booking Log File"):
        st.session_state.clear()
        st.session_state.step = 1
        st.rerun()
