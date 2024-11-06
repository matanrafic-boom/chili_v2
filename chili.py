import streamlit as st
import json
import pandas as pd
import requests
from collections import Counter, defaultdict
import os
from dotenv import load_dotenv
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# Set page config (this must be the first Streamlit command)
st.set_page_config(page_title="BizOps 💥", page_icon="💥", layout="wide")

# Load environment variables from .env file
load_dotenv()

# Chili Piper API configuration
API_BASE_URL = "https://edge.na.chilipiper.com"

# Try to get API_KEY from environment variable first, then from Streamlit secrets
API_KEY = os.getenv("CHILI_API_KEY")
if not API_KEY:
    try:
        API_KEY = st.secrets["CHILI_API_KEY"]
    except FileNotFoundError:
        st.error("CHILI_API_KEY is not set. Please set it in your .env file or Streamlit secrets.")
        st.stop()

if not API_KEY:
    st.error("CHILI_API_KEY is not set. Please set it in your .env file or Streamlit secrets.")
    st.stop()

# Workspace ID to name mapping
WORKSPACE_NAMES = {
    "64ad3cc865a4906cd3cc2dcf": "Sales",
    "61b9daad2747672e7282273d": "CS"
}

# Custom CSS
st.markdown("""
    <style>
    .main {
        background-color: #f0f2f6;
    }
    .stApp {
        margin: 0 auto;
    }
    h1 {
        color: #1E3A8A;
        text-align: center;
        padding: 20px 0;
    }
    h2 {
        color: #2563EB;
        border-bottom: 2px solid #2563EB;
        padding-bottom: 10px;
    }
    .stMetric {
        background-color: white;
        padding: 15px;
        border-radius: 5px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .stTable {
        background-color: white;
        border-radius: 5px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .sidebar .sidebar-content {
        background-color: #f8fafc;
    }
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #f0f2f6;
        color: #4B5563;
        text-align: center;
        padding: 10px 0;
        font-size: 0.8em;
    }
    
    /* Updated table styles for classic Microsoft look */
    .table-container {
        background: white;
        border: 1px solid #c6c6c6;
        border-collapse: collapse;
        width: 100%;
    }
    
    .table-header {
        background-color: #f0f0f0;
        border-bottom: 1px solid #c6c6c6;
        padding: 8px;
        font-weight: bold;
    }
    
    .table-row {
        border-bottom: 1px solid #e0e0e0;
        background-color: white;
    }
    
    .table-row:nth-child(even) {
        background-color: #f9f9f9;
    }
    
    .table-row:hover {
        background-color: #f5f5f5;
    }
    
    .sort-button {
        background-color: #f0f0f0;
        border: none;
        color: #000000;
        text-align: left;
        width: 100%;
        padding: 8px;
        font-weight: bold;
    }
    
    .sort-button:hover {
        background-color: #e3e3e3;
    }
    
    .action-button {
        background-color: #f0f0f0;
        border: 1px solid #c6c6c6;
        padding: 4px 8px;
        margin: 2px;
        cursor: pointer;
    }
    
    .action-button:hover {
        background-color: #e3e3e3;
    }
    
    /* New queue header styles */
    .queue-header {
        background-color: #f8f9fa;
        padding: 15px 20px;
        border-radius: 8px;
        margin-bottom: 10px;
        border: 1px solid #e9ecef;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    
    .queue-title {
        font-size: 18px;
        color: #1a1a1a;
        margin: 0;
        font-weight: 500;
    }
    
    .queue-link {
        color: #2563EB;
        text-decoration: none;
        padding: 5px 10px;
        border-radius: 4px;
        font-size: 14px;
    }
    
    .queue-link:hover {
        background-color: #f0f0f0;
    }
    
    /* Improved table styles */
    .table-container {
        background: white;
        border: 1px solid #dee2e6;
        border-radius: 6px;
        margin-top: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    
    .table-header {
        background-color: #f8f9fa;
        padding: 12px 15px;
        border-bottom: 2px solid #dee2e6;
        font-weight: 500;
    }
    
    .table-row {
        padding: 10px 15px;
        border-bottom: 1px solid #dee2e6;
        transition: background-color 0.2s;
    }
    
    .table-row:last-child {
        border-bottom: none;
    }
    
    .table-row:hover {
        background-color: #f8f9fa;
    }
    
    /* Button styles */
    .action-button {
        padding: 4px 12px;
        border-radius: 4px;
        border: 1px solid #dee2e6;
        background-color: white;
        color: #1a1a1a;
        cursor: pointer;
        transition: all 0.2s;
    }
    
    .action-button:hover {
        background-color: #f8f9fa;
        border-color: #c6c6c6;
    }
    
    .edit-button {
        color: #2563EB;
        border-color: #2563EB;
    }
    
    .remove-button {
        color: #dc3545;
        border-color: #dc3545;
    }
    
    /* Add new rep button */
    .add-rep-button {
        margin-bottom: 15px;
        padding: 8px 16px;
        background-color: #2563EB;
        color: white;
        border: none;
        border-radius: 4px;
        cursor: pointer;
    }
    
    .add-rep-button:hover {
        background-color: #1d4ed8;
    }
    
    /* Form styles */
    .edit-form {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 6px;
        margin: 10px 0;
        border: 1px solid #dee2e6;
    }
    
    .form-group {
        margin-bottom: 15px;
    }
    
    .form-label {
        font-weight: 500;
        margin-bottom: 5px;
        display: block;
    }
    </style>
    """, unsafe_allow_html=True)

# Add this at the top level with other constants
BIZOPS_COMPLIMENTS = [
    "BizOps rocks! 🚀",
    "Best team ever! ⭐",
    "You make data beautiful! 📊",
    "BizOps saves the day! 💪",
    "Automation wizards! ✨"
]

# Add these constants at the top with other constants
AUDIT_SPREADSHEET_URL = "YOUR_GOOGLE_SHEET_URL"  # You'll need to create this and share it
ACTION_TYPES = {
    "WEIGHT_UPDATE": "Updated weight",
    "REP_REMOVE": "Removed rep from queue",
    "REP_ADD": "Added rep to queue"
}

# Add these constants at the top of the file with other constants
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Function to fetch queue data from Chili Piper API
def fetch_queue_data():
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    params = {
        "pageSize": "100"
    }
    try:
        logger.info("Fetching queue data from API")
        response = requests.get(f"{API_BASE_URL}/queue", headers=headers, params=params)
        response.raise_for_status()
        logger.info(f"Successfully fetched queue data. Status code: {response.status_code}")
        return response.json()
    except Exception as e:
        logger.error(f"Error fetching queue data: {str(e)}")
        raise

# Function to generate statistics
def generate_statistics(json_data, selected_workspace, selected_rep, selected_size):
    # Filter active queues with members and sort by name, excluding "Existing Customer - Owner"
    active_queues = sorted(
        [q for q in json_data['elements'] if q['active'] and q.get('members') and 
         q['name'] != "Existing Customer - Owner" and
         (selected_workspace == "All" or WORKSPACE_NAMES.get(q['workspaceId'], q['workspaceId']) == selected_workspace) and
         (selected_size == "All" or extract_size_range(q) == selected_size)],
        key=lambda x: len(x.get('members', [])),
        reverse=True
    )

    stats = {
        'total_queues': len(active_queues),
        'total_reps': sum(len(q.get('members', [])) for q in active_queues),
        'queues_by_size': Counter(len(q.get('members', [])) for q in active_queues),
        'reps_by_queue': {q['name']: len(q.get('members', [])) for q in active_queues},
        'main_reps': sum(1 for q in active_queues for m in q.get('members', []) if m.get('main', False)),
        'mandatory_reps': sum(1 for q in active_queues for m in q.get('members', []) if m.get('mandatory', False)),
        'queue_pivot': {},
        'rep_pivot': defaultdict(list),
        'workspaces': set(),
        'queue_links': {}
    }

    # Generate queue pivot and rep pivot
    for queue in active_queues:
        queue_name = queue['name']
        stats['workspaces'].add(WORKSPACE_NAMES.get(queue['workspaceId'], queue['workspaceId']))
        members = [(member['name'], member['weight'], member['order'], member['initialOrder'], member['id'], queue['id']) 
                  for member in queue.get('members', [])
                  if selected_rep == "All" or member['name'] == selected_rep]
        if members:
            stats['queue_pivot'][queue_name] = sorted(members, key=lambda x: x[2])
        
        for member in queue.get('members', []):
            if selected_rep == "All" or member['name'] == selected_rep:
                stats['rep_pivot'][member['name']].append((
                    queue_name, 
                    member['weight'], 
                    member['order'], 
                    member['initialOrder'],
                    member['id'],  # Add user ID
                    queue['id']    # Add queue ID
                ))
        
        # Sort rep_pivot entries by order
        for rep_name in stats['rep_pivot']:
            stats['rep_pivot'][rep_name] = sorted(stats['rep_pivot'][rep_name], key=lambda x: x[2])

        # Add queue link
        stats['queue_links'][queue_name] = f"https://connecteam.na.chilipiper.com/admin-center/meetings/{queue['workspaceId']}/queues/edit/{queue['id']}"

    # New: Generate AE participation data
    ae_participation = defaultdict(lambda: defaultdict(int))
    sales_queues = []
    cs_queues = []

    for queue in active_queues:
        queue_name = queue['name']
        workspace = WORKSPACE_NAMES.get(queue['workspaceId'], queue['workspaceId'])
        
        if workspace == "Sales":
            sales_queues.append(queue_name)
        elif workspace == "CS":
            cs_queues.append(queue_name)
        
        for member in queue.get('members', []):
            ae_participation[member['name']][queue_name] = member['weight']

    # Remove users that only appear in "Existing Customer - Owner" queue
    existing_customer_owner_queue = next((q for q in json_data['elements'] if q['name'] == "Existing Customer - Owner"), None)
    if existing_customer_owner_queue:
        existing_customer_owner_members = set(m['name'] for m in existing_customer_owner_queue.get('members', []))
        users_to_remove = existing_customer_owner_members - set(ae_participation.keys())
        for user in users_to_remove:
            if user in ae_participation:
                del ae_participation[user]

    stats['ae_participation'] = ae_participation
    stats['sales_queues'] = sorted(sales_queues)
    stats['cs_queues'] = sorted(cs_queues)

    # Inside the generate_statistics function, add this before returning stats:
    cs_users = set()
    for queue in active_queues:
        if WORKSPACE_NAMES.get(queue['workspaceId']) == "CS":
            for member in queue.get('members', []):
                cs_users.add(member['name'])
    stats['cs_users'] = list(cs_users)

    return stats

# Add this new function to handle API calls
def update_queue_member_weight(queue_id, user_ids, weight, queue_name, rep_name):
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    payload = {
        "users": user_ids,
        "weight": weight
    }
    try:
        response = requests.post(
            f"{API_BASE_URL}/queue/{queue_id}/user/update/weighted",
            headers=headers,
            json=payload
        )
        response.raise_for_status()
        
        # Log the successful action
        log_action(
            action_type=ACTION_TYPES["WEIGHT_UPDATE"],
            queue_name=queue_name,
            rep_name=rep_name,
            details=f"Updated weight to {weight}"
        )
        return True
    except Exception as e:
        st.error(f"Error updating weights: {str(e)}")
        return False

# Add these new functions to handle the API calls

def remove_reps_from_queue(queue_id, user_ids):
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    try:
        logger.info(f"Attempting to remove users from queue_id={queue_id}, user_ids={user_ids}")
        response = requests.post(
            f"{API_BASE_URL}/queue/{queue_id}/user/unassign",
            headers=headers,
            json=user_ids
        )
        response.raise_for_status()
        logger.info(f"Successfully removed users from queue_id={queue_id}")
        return True
    except Exception as e:
        logger.error(f"Error removing users: queue_id={queue_id}, error={str(e)}")
        st.error(f"Error removing users: {str(e)}")
        return False

def add_rep_to_queue(queue_id, user_ids, weight=None):
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    
    try:
        if weight:
            logger.info(f"Attempting to add users with weight to queue_id={queue_id}, user_ids={user_ids}, weight={weight}")
            response = requests.post(
                f"{API_BASE_URL}/queue/{queue_id}/user/assign/weighted",
                headers=headers,
                json={"users": user_ids, "weight": weight}
            )
        else:
            logger.info(f"Attempting to add users to queue_id={queue_id}, user_ids={user_ids}")
            response = requests.post(
                f"{API_BASE_URL}/queue/{queue_id}/user/assign",
                headers=headers,
                json=user_ids
            )
        
        response.raise_for_status()
        logger.info(f"Successfully added users to queue_id={queue_id}")
        return True
    except Exception as e:
        logger.error(f"Error adding users: queue_id={queue_id}, error={str(e)}")
        st.error(f"Error adding users: {str(e)}")
        return False

# Add this new function to handle the delete confirmation UI
def show_delete_confirmation(queue_name, rep_name, queue_id, user_id, i, index):
    with st.form(key=f"delete_confirmation_{i}_{index}"):
        st.warning(f"⚠️ Are you sure you want to remove {rep_name} from {queue_name}?")
        st.write("This action cannot be undone.")
        
        # Replace text input with selectbox
        confirmation = st.selectbox(
            "Select a compliment to confirm removal:",
            options=BIZOPS_COMPLIMENTS,
            key=f"delete_confirmation_input_{i}_{index}"
        )
        
        col1, col2 = st.columns([1, 1])
        with col1:
            confirm_button = st.form_submit_button(
                "Confirm Removal",
                type="primary",
                help="This will permanently remove the rep from this queue"
            )
        with col2:
            cancel_button = st.form_submit_button(
                "Cancel",
                help="Cancel the removal process"
            )
        
        if cancel_button:
            st.session_state[f'removing_rep_{i}_{index}'] = False
            st.rerun()
            
        if confirm_button:
            if confirmation in BIZOPS_COMPLIMENTS:  # Check if a compliment was selected
                if remove_reps_from_queue(queue_id, [user_id]):
                    st.success(f"Successfully removed {rep_name} from the queue")
                    st.session_state[f'removing_rep_{i}_{index}'] = False
                    st.rerun()
            else:
                st.error("Please select a compliment to confirm the removal")

# Add this new function to validate the add rep form
def validate_add_rep_form(weight, order, main, mandatory, lock):
    if not weight or weight < 1 or weight > 100:
        return False, "Weight must be between 1 and 100"
    if order < 0:
        return False, "Order must be non-negative"
    return True, ""

# Add this new function at the top level
def get_sort_key(i, queue_name):
    """Generate a unique key for storing sort state"""
    return f"sort_state_{queue_name}_{i}"

# Modify the extract_size_range function:
def extract_size_range(queue):
    for rule in queue.get('rules', []):
        if (rule.get('entity') == 'Contact' and 
            rule.get('field') == 'numofemployeesrange' and 
            rule.get('operator') == '='):
            value = rule.get('value')
            # Group the ranges
            if value:
                try:
                    # Extract first number from range (e.g., "11-30" -> 11)
                    first_number = int(value.split('-')[0])
                    if first_number <= 50:
                        return "1-50"
                    elif first_number <= 100:
                        return "51-100"
                    else:
                        return "101 and above"
                except:
                    return "No Size"
    return "No Size"  # Return "No Size" if no size rule is found

# Add this new function to handle order updates
def update_queue_member_order(queue_id, user_id, order):
    # Note: Since there's no direct API for order updates, we'll use the weight update for now
    # This is a placeholder for when the order update API becomes available
    return True

# Add this new function to handle audit logging
def log_action(action_type: str, queue_name: str, rep_name: str, details: str):
    try:
        sheet = get_google_sheet()
        if sheet:
            # Get user email from Streamlit's authentication
            user_email = "unknown"
            try:
                if st.runtime.exists():
                    user_email = st.session_state.get('user_email', 'unknown')
                    if user_email == 'unknown' and hasattr(st, 'experimental_user'):
                        user_email = st.experimental_user.email
                        st.session_state['user_email'] = user_email
            except Exception as e:
                logger.error(f"Error getting user email: {str(e)}")

            # Create the log entry with user email
            log_entry = [
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                user_email,
                action_type,
                queue_name,
                rep_name,
                details
            ]
            
            # Log the user info for debugging
            logger.info(f"Logging action by user: {user_email}")
            
            # Append to the Google Sheet
            sheet.append_row(log_entry)
            logger.info(f"Successfully logged action: {action_type} by {user_email}")
        else:
            logger.error("Failed to get Google Sheet connection")
            
    except Exception as e:
        logger.error(f"Failed to log action: {str(e)}")

# Add this function to handle the Google Sheets connection
def get_google_sheet():
    try:
        credentials = ServiceAccountCredentials.from_json_keyfile_dict(
            st.secrets["gcp_service_account"], 
            SCOPES
        )
        gc = gspread.authorize(credentials)
        return gc.open_by_url(st.secrets["private"]["sheet_url"]).sheet1
    except Exception as e:
        logger.error(f"Failed to connect to Google Sheets: {str(e)}")
        return None

# Streamlit app
def main():
    st.title('BizOps 💥')
    st.subheader('Sales Queue Statistics Dashboard')

    # Fetch queue data from API
    try:
        json_data = fetch_queue_data()
        all_workspaces = set(WORKSPACE_NAMES.get(q['workspaceId'], q['workspaceId']) for q in json_data['elements'])
        all_reps = set(member['name'] for q in json_data['elements'] for member in q.get('members', []))
        
        # Define the fixed size ranges in the desired order, including "No Size"
        all_size_ranges = ["1-50", "51-100", "101 and above", "No Size"]
        
        # Verify which ranges have active queues
        active_size_ranges = set()
        for queue in json_data['elements']:
            if queue['active'] and queue.get('members'):
                size_range = extract_size_range(queue)
                if size_range:
                    active_size_ranges.add(size_range)
        
        # Only show ranges that have active queues
        all_size_ranges = [size for size in all_size_ranges if size in active_size_ranges]
        
    except Exception as e:
        st.error(f"Error fetching data from Chili Piper API: {str(e)}")
        return

    # Filters
    col1, col2, col3 = st.columns(3)
    with col1:
        # Get list of workspaces and ensure "Sales" is the default
        workspace_options = ["All"] + list(all_workspaces)
        default_index = workspace_options.index("Sales") if "Sales" in workspace_options else 0
        
        selected_workspace = st.selectbox(
            "**Select Workspace**", 
            workspace_options,
            index=default_index  # This will select "Sales" by default
        )
        
        # Force "Sales" as default if not already selected
        if "workspace_initialized" not in st.session_state:
            selected_workspace = "Sales"
            st.session_state.workspace_initialized = True
            st.rerun()  # Rerun to apply the default selection

    with col2:
        selected_size = st.selectbox(
            "**Select Size Range**",
            ["All"] + all_size_ranges,
            help="Filter queues by employee size range"
        )
    with col3:
        selected_rep = st.selectbox(
            "**Select Rep**", 
            ["All"] + list(all_reps), 
            key="rep_selector"
        )

    # Generate statistics based on all filters
    stats = generate_statistics(json_data, selected_workspace, selected_rep, selected_size)

    # Sidebar navigation
    st.sidebar.title("Navigation")
    sections = [
        "Queues and Reps",
        "Employees and Their Queues",
        "AE Participation",
        "Reps by Queue",
        "Queues by Size",
        "Overall Statistics",
        "Audit Log"
    ]
    for section in sections:
        if st.sidebar.button(section):
            st.query_params["section"] = section.lower().replace(" ", "_")

    # Get the current section from query parameters
    current_section = st.query_params.get("section", "queues_and_reps")

    # If a specific rep is selected, switch to "Employees and Their Queues" section
    if selected_rep != "All":
        current_section = "employees_and_their_queues"
        st.query_params["section"] = current_section

    # Display queues and reps in table view
    if current_section == "queues_and_reps":
        st.header('Queues and Reps')
        for i, (queue_name, reps) in enumerate(stats['queue_pivot'].items()):
            # Queue header with improved styling
            st.markdown(f"""
                <div class="queue-header">
                    <h3 class="queue-title">{queue_name} ({len(reps)} reps)</h3>
                    <a href="{stats['queue_links'][queue_name]}" target="_blank" class="queue-link">View in Chili Piper →</a>
                </div>
            """, unsafe_allow_html=True)
            
            with st.expander("Show Details", expanded=False):
                # Initialize sort state
                sort_key = get_sort_key(i, queue_name)
                if sort_key not in st.session_state:
                    st.session_state[sort_key] = {'column': 'Order', 'ascending': True}

                # Create DataFrame first
                rep_df = pd.DataFrame(reps, columns=['Rep Name', 'Weight', 'Order', 'Initial Order', 'User ID', 'Queue ID'])

                # Add New Rep button
                if st.button("➕ Add New Rep", key=f"add_rep_button_{i}", 
                            help="Add a new rep to this queue",
                            type="primary"):
                    st.session_state[f'adding_rep_{i}'] = True
                
                # Add New Rep form
                if st.session_state.get(f'adding_rep_{i}', False):
                    with st.form(key=f"add_rep_form_{i}"):
                        st.subheader(f"Add New Reps to {queue_name}")
                        
                        # Get all available reps that aren't in this queue
                        current_rep_names = {rep[0] for rep in reps}
                        available_reps = list(all_reps - current_rep_names)
                        
                        # Multiple Rep selection
                        selected_new_reps = st.multiselect(
                            "Select Reps *",
                            options=available_reps,
                            help="Select one or more reps to add to this queue"
                        )
                        
                        # Single weight input for all selected reps
                        new_weight = st.number_input(
                            "Weight for all selected reps *",
                            min_value=1,
                            max_value=100,
                            value=50,
                            help="Weight determines routing priority (1-100). Will be applied to all selected reps."
                        )
                        
                        # Action buttons
                        col1, col2 = st.columns([1, 1])
                        with col1:
                            if st.form_submit_button("Add Reps"):
                                success = True
                                failed_reps = []
                                
                                for selected_rep in selected_new_reps:
                                    # Find user ID for selected rep
                                    user_id = next((member['id'] for q in json_data['elements'] 
                                                  for member in q.get('members', []) 
                                                  if member['name'] == selected_rep), None)
                                    
                                    if user_id:
                                        if not add_rep_to_queue(queue_id, [user_id], new_weight):
                                            success = False
                                            failed_reps.append(selected_rep)
                                    else:
                                        success = False
                                        failed_reps.append(selected_rep)
                                
                                if success:
                                    st.success(f"""
                                        ✅ Successfully added {len(selected_new_reps)} reps to {queue_name}:
                                        - Reps: {', '.join(selected_new_reps)}
                                        - Weight: {new_weight}
                                    """)
                                    st.session_state[f'adding_rep_{i}'] = False
                                    st.rerun()
                                else:
                                    if failed_reps:
                                        st.error(f"Failed to add some reps: {', '.join(failed_reps)}")
                                    else:
                                        st.error("Failed to add reps. Please try again.")
                        with col2:
                            if st.form_submit_button("Cancel"):
                                st.session_state[f'adding_rep_{i}'] = False
                                st.rerun()

                # Create table container
                st.markdown('<div class="table-container">', unsafe_allow_html=True)
                
                # Table header
                header_cols = st.columns([3, 2, 2, 1.5, 1.5])
                with header_cols[0]:
                    if st.button("Rep Name", key=f"sort_name_{i}", 
                               help="Click to sort by name"):
                        if st.session_state[sort_key]['column'] == 'Rep Name':
                            st.session_state[sort_key]['ascending'] = not st.session_state[sort_key]['ascending']
                        else:
                            st.session_state[sort_key] = {'column': 'Rep Name', 'ascending': True}
                with header_cols[1]:
                    if st.button("Weight", key=f"sort_weight_{i}",
                               help="Click to sort by weight"):
                        if st.session_state[sort_key]['column'] == 'Weight':
                            st.session_state[sort_key]['ascending'] = not st.session_state[sort_key]['ascending']
                        else:
                            st.session_state[sort_key] = {'column': 'Weight', 'ascending': True}
                with header_cols[2]:
                    if st.button("Order", key=f"sort_order_{i}",
                               help="Click to sort by order"):
                        if st.session_state[sort_key]['column'] == 'Order':
                            st.session_state[sort_key]['ascending'] = not st.session_state[sort_key]['ascending']
                        else:
                            st.session_state[sort_key] = {'column': 'Order', 'ascending': True}
                with header_cols[3]:
                    st.write("Edit")
                with header_cols[4]:
                    st.write("Remove")

                st.markdown('<hr style="margin: 0; padding: 0; border-color: #eee;">', unsafe_allow_html=True)

                # Sort the DataFrame
                if st.session_state[sort_key]['column']:
                    rep_df = rep_df.sort_values(
                        by=st.session_state[sort_key]['column'],
                        ascending=st.session_state[sort_key]['ascending']
                    )

                # Display sort indicator
                if st.session_state[sort_key]['column']:
                    st.caption(f"Sorted by {st.session_state[sort_key]['column']} "
                             f"({'ascending' if st.session_state[sort_key]['ascending'] else 'descending'})")

                # Display rep data
                for index, row in rep_df.iterrows():
                    cols = st.columns([3, 2, 2, 1.5, 1.5])
                    with cols[0]:
                        st.write(f"{row['Rep Name']}")
                    with cols[1]:
                        st.write(f"{row['Weight']}")
                    with cols[2]:
                        st.write(f"{row['Order']}")
                    with cols[3]:
                        if st.button("Edit", key=f"edit_button_{i}_{index}",
                                   help="Edit this rep's settings"):
                            st.session_state[f'editing_{i}_{index}'] = True
                    with cols[4]:
                        if st.button("Remove", key=f"remove_button_{i}_{index}",
                                   type="secondary",
                                   help="Remove this rep from the queue"):
                            st.session_state[f'removing_rep_{i}_{index}'] = True

                    # Show edit form if editing is active
                    if st.session_state.get(f'editing_{i}_{index}', False):
                        with st.form(key=f"edit_form_{i}_{index}"):
                            new_weight = st.number_input(
                                "New Weight",
                                min_value=1,
                                max_value=100,
                                value=int(row['Weight'])
                            )
                            new_order = st.number_input(
                                "New Order",
                                min_value=0,
                                max_value=len(reps)-1,
                                value=int(row['Order'])
                            )
                            col1, col2 = st.columns([1, 1])
                            with col1:
                                if st.form_submit_button("Save"):
                                    if update_queue_member_weight(row['Queue ID'], [row['User ID']], new_weight, queue_name, row['Rep Name']):
                                        st.success(f"Successfully updated {row['Rep Name']}'s weight to {new_weight}")
                                        st.session_state[f'editing_{i}_{index}'] = False
                                        st.rerun()
                            with col2:
                                if st.form_submit_button("Cancel"):
                                    st.session_state[f'editing_{i}_{index}'] = False
                                    st.rerun()

                    # Show delete confirmation if remove button was clicked
                    if st.session_state.get(f'removing_rep_{i}_{index}', False):
                        show_delete_confirmation(
                            queue_name,
                            row['Rep Name'],
                            row['Queue ID'],
                            row['User ID'],
                            i,
                            index
                        )

    # Display employees and their queues
    if current_section == "employees_and_their_queues":
        st.header('Employees and Their Queues')
        
        # Sort the rep_pivot dictionary by the number of queues (in descending order)
        sorted_reps = sorted(stats['rep_pivot'].items(), key=lambda x: len(x[1]), reverse=True)
        
        for i, (employee_name, queues) in enumerate(sorted_reps):
            with st.expander(f"{employee_name} ({len(queues)} queues)", 
                           expanded=(employee_name == selected_rep and selected_rep != "All")):
                
                # Add New Queue button
                if st.button("➕ Add to Queue", key=f"add_queue_button_{i}"):
                    st.session_state[f'adding_queue_{i}'] = True

                # Add to Queue form
                if st.session_state.get(f'adding_queue_{i}', False):
                    with st.form(key=f"add_queue_form_{i}"):
                        st.subheader(f"Add {employee_name} to Queue(s)")
                        
                        # Get all available queues for the workspace
                        current_queue_names = {q[0] for q in queues}
                        available_queues = [
                            q for q in json_data['elements'] 
                            if q['active'] and 
                            q['name'] not in current_queue_names and
                            WORKSPACE_NAMES.get(q['workspaceId'], '') == selected_workspace
                        ]
                        
                        # Multiple Queue selection
                        selected_queues = st.multiselect(
                            "Select Queue(s) *",
                            options=[q['name'] for q in available_queues],
                            help="Select one or more queues to add the rep to"
                        )
                        
                        # Single weight input for all selected queues
                        new_weight = st.number_input(
                            "Weight for all selected queues *",
                            min_value=1,
                            max_value=100,
                            value=50,
                            help="Weight determines routing priority (1-100). Will be applied to all selected queues."
                        )
                        
                        # Action buttons
                        col1, col2 = st.columns([1, 1])
                        with col1:
                            if st.form_submit_button("Add to Queues"):
                                success = True
                                failed_queues = []
                                
                                # Get user ID from the existing queues data
                                user_id = next((q[4] for q in queues), None)  # Index 4 is User ID in the tuple
                                
                                if user_id:
                                    for queue_name in selected_queues:
                                        queue_obj = next((q for q in available_queues if q['name'] == queue_name), None)
                                        if queue_obj:
                                            if not add_rep_to_queue(queue_obj['id'], [user_id], new_weight):
                                                success = False
                                                failed_queues.append(queue_name)
                                    
                                    if success:
                                        st.success(f"Successfully added {employee_name} to all selected queues with weight {new_weight}")
                                        st.session_state[f'adding_queue_{i}'] = False
                                        st.rerun()
                                    else:
                                        st.error(f"Failed to add to some queues: {', '.join(failed_queues)}")
                                else:
                                    st.error("Could not find user ID")
                        with col2:
                            if st.form_submit_button("Cancel"):
                                st.session_state[f'adding_queue_{i}'] = False
                                st.rerun()

                # Display existing queues
                queue_df = pd.DataFrame(queues, columns=[
                    'Queue Name', 'Weight', 'Order', 'Initial Order', 'User ID', 'Queue ID'
                ])
                
                # Display each queue as a row with actions (keep only the columns we want to show)
                for index, row in queue_df.iterrows():
                    cols = st.columns([4, 3, 1.5, 1.5])
                    with cols[0]:
                        st.write(f"{row['Queue Name']}")
                    with cols[1]:
                        st.write(f"Weight: {row['Weight']}")
                    with cols[2]:
                        if st.button("Edit", key=f"edit_button_{i}_{index}",
                                   help="Edit rep's weight in this queue"):
                            st.session_state[f'editing_{i}_{index}'] = True
                    with cols[3]:
                        if st.button("Remove", key=f"remove_button_{i}_{index}",
                                   type="secondary",
                                   help="Remove rep from this queue"):
                            st.session_state[f'removing_rep_{i}_{index}'] = True

                    # Show edit form if editing is active
                    if st.session_state.get(f'editing_{i}_{index}', False):
                        with st.form(key=f"edit_form_{i}_{index}"):
                            new_weight = st.number_input(
                                "New Weight",
                                min_value=1,
                                max_value=100,
                                value=int(row['Weight'])
                            )
                            col1, col2 = st.columns([1, 1])
                            with col1:
                                if st.form_submit_button("Save"):
                                    if update_queue_member_weight(row['Queue ID'], [row['User ID']], new_weight, queue_name, row['Rep Name']):
                                        st.success(f"""
                                            ✅ Successfully updated {employee_name}'s weight in {row['Queue Name']}:
                                            - Weight: {row['Weight']} → {new_weight}
                                        """)
                                        st.session_state[f'editing_{i}_{index}'] = False
                                        st.rerun()
                                    else:
                                        st.error("Failed to update weight. Please try again.")
                            with col2:
                                if st.form_submit_button("Cancel"):
                                    st.session_state[f'editing_{i}_{index}'] = False
                                    st.rerun()

                    # Show delete confirmation if remove button was clicked
                    if st.session_state.get(f'removing_rep_{i}_{index}', False):
                        with st.form(key=f"delete_confirmation_{i}_{index}"):
                            st.warning(f"⚠️ Are you sure you want to remove {employee_name} from {row['Queue Name']}?")
                            st.write("This action cannot be undone.")
                            
                            # Replace text input with selectbox
                            confirmation = st.selectbox(
                                "Select a compliment to confirm removal:",
                                options=BIZOPS_COMPLIMENTS,
                                key=f"delete_confirmation_input_{i}_{index}"
                            )
                            
                            col1, col2 = st.columns([1, 1])
                            with col1:
                                confirm_button = st.form_submit_button(
                                    "Confirm Removal",
                                    type="primary",
                                    help="This will permanently remove the rep from this queue"
                                )
                            with col2:
                                cancel_button = st.form_submit_button(
                                    "Cancel",
                                    help="Cancel the removal process"
                                )
                            
                            if cancel_button:
                                st.session_state[f'removing_rep_{i}_{index}'] = False
                                st.rerun()
                                
                            if confirm_button:
                                if confirmation in BIZOPS_COMPLIMENTS:  # Check if a compliment was selected
                                    if remove_reps_from_queue(row['Queue ID'], [row['User ID']]):
                                        st.success(f"Successfully removed {employee_name} from {row['Queue Name']}")
                                        st.session_state[f'removing_rep_{i}_{index}'] = False
                                        st.rerun()
                                else:
                                    st.error("Please select a compliment to confirm the removal")

    # Reps by Queue
    if current_section == "reps_by_queue":
        st.header('Reps by Queue')
        reps_by_queue_df = pd.DataFrame.from_dict(stats['reps_by_queue'], orient='index', columns=['Count']).reset_index()
        reps_by_queue_df.columns = ['Queue Name', 'Count']
        reps_by_queue_df = reps_by_queue_df.sort_values('Count', ascending=False)
        fig = px.bar(reps_by_queue_df, x='Queue Name', y='Count', title='Reps by Queue')
        fig.update_layout(xaxis_tickangle=-45, height=600)
        st.plotly_chart(fig, use_container_width=True, key="reps_by_queue_chart")

    # Queues by Size
    if current_section == "queues_by_size":
        st.header('Queues by Size')
        queues_by_size_df = pd.DataFrame.from_dict(stats['queues_by_size'], orient='index', columns=['Count']).reset_index()
        queues_by_size_df.columns = ['Number of Reps', 'Count']
        fig = px.pie(queues_by_size_df, values='Count', names='Number of Reps', title='Distribution of Queue Sizes')
        st.plotly_chart(fig, use_container_width=True, key="queues_by_size_chart")

    # Display overall statistics
    if current_section == "overall_statistics":
        st.header('Overall Statistics')
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Active Queues", stats['total_queues'])
        col2.metric("Total Reps", stats['total_reps'])
        col3.metric("Main Reps", stats['main_reps'])
        col4.metric("Mandatory Reps", stats['mandatory_reps'])

    # New: AE Participation section
    if current_section == "ae_participation":
        st.header('AE Participation in Queues')

        # CSS for the scrollable table with fixed header and left column
        st.markdown("""
        <style>
        .scrollable-table-container {
            max-height: 500px;
            overflow: auto;
            border: 1px solid #ddd;
            position: relative;
        }
        .scrollable-table {
            border-collapse: separate;
            border-spacing: 0;
        }
        .scrollable-table th, .scrollable-table td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: center;
        }
        .scrollable-table thead th {
            position: sticky;
            top: 0;
            background-color: #f1f1f1;
            z-index: 2;
        }
        .scrollable-table td:first-child, .scrollable-table th:first-child {
            position: sticky;
            left: 0;
            background-color: #f1f1f1;
            z-index: 1;
        }
        .scrollable-table thead th:first-child {
            z-index: 3;
        }
        .scrollable-table tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        .scrollable-table tr:hover {
            background-color: #f5f5f5;
        }
        .zero-value {
            background-color: #d3d3d3;
        }
        .queue-link {
            color: #2563EB;
            text-decoration: none;
        }
        .queue-link:hover {
            text-decoration: underline;
        }
        </style>
        """, unsafe_allow_html=True)

        def create_scrollable_table(df, queue_links):
            df = df.applymap(lambda x: f"{x:.0f}")  # Remove decimal places
            table_html = f"""
            <div class="scrollable-table-container">
                <table class="scrollable-table">
                    <thead>
                        <tr>
                            <th>Rep Name</th>
                            {''.join(f'<th><a href="{queue_links.get(col, "#")}" class="queue-link" target="_blank">{col}</a></th>' for col in df.columns)}
                        </tr>
                    </thead>
                    <tbody>
                        {''.join(f'<tr><td>{index}</td>' + ''.join(f'<td class="{"zero-value" if cell == "0" else ""}">{cell}</td>' for cell in row) + '</tr>' for index, row in df.iterrows())}
                    </tbody>
                </table>
            </div>
            """
            return table_html

        # Sales Queues
        st.subheader('Sales Queues')
        if stats['sales_queues']:
            sales_df = pd.DataFrame(stats['ae_participation']).T[stats['sales_queues']]
            sales_df = sales_df.fillna(0)  # Replace NaN with 0
            # Filter out reps that don't appear in any sales queue
            sales_reps = list(sales_df.index[sales_df.sum(axis=1) > 0])
            sales_df = sales_df.loc[sales_reps]
            # Sort columns based on the number of reps in each queue
            sales_df = sales_df.loc[:, sales_df.sum().sort_values(ascending=False).index]
            st.markdown(create_scrollable_table(sales_df, stats['queue_links']), unsafe_allow_html=True)
        else:
            st.write("No Sales Queues found.")

        # CS Queues
        st.subheader('CS Queues')
        if stats['cs_queues']:
            cs_df = pd.DataFrame(stats['ae_participation']).T[stats['cs_queues']]
            cs_df = cs_df.fillna(0)  # Replace NaN with 0
            # Filter to show only reps that are in CS workspace
            cs_df = cs_df.loc[cs_df.index.intersection(stats['cs_users'])]
            # Sort columns based on the number of reps in each queue
            cs_df = cs_df.loc[:, cs_df.sum().sort_values(ascending=False).index]
            st.markdown(create_scrollable_table(cs_df, stats['queue_links']), unsafe_allow_html=True)
        else:
            st.write("No CS Queues found.")

        # Display CS Users
        st.subheader('CS Users')
        st.write(f"Total CS Users: {len(stats['cs_users'])}")
        st.write("CS Users active in any queue:")
        st.write(", ".join(sorted(stats['cs_users'])))

        # Debug Information
        st.subheader('Debug Information')
        st.write(f"Number of CS Queues: {len(stats['cs_queues'])}")
        st.write("CS Queues:")
        st.write(stats['cs_queues'])

    # New: Audit Log section
    if current_section == "audit_log":
        st.header("Audit Log")
        
        # Add filters
        col1, col2, col3 = st.columns(3)
        with col1:
            date_filter = st.date_input("Filter by Date")
        with col2:
            action_filter = st.selectbox(
                "Filter by Action",
                ["All"] + list(ACTION_TYPES.values())
            )
        with col3:
            user_filter = st.text_input("Filter by User")
        
        try:
            # Get the audit log data
            sheet = get_google_sheet()
            if sheet:
                # Get all records
                records = sheet.get_all_records()
                df = pd.DataFrame(records)
                
                # Apply filters
                if date_filter:
                    df['Date'] = pd.to_datetime(df['Timestamp']).dt.date
                    df = df[df['Date'] == date_filter]
                
                if action_filter != "All":
                    df = df[df['Action'] == action_filter]
                    
                if user_filter:
                    df = df[df['User'].str.contains(user_filter, case=False)]
                
                # Display the filtered log
                st.dataframe(
                    df,
                    column_config={
                        'Timestamp': st.column_config.DatetimeColumn(
                            'Time',
                            format="DD/MM/YY HH:mm:ss"
                        ),
                        'Details': st.column_config.TextColumn(
                            'Details',
                            width='large'
                        )
                    },
                    hide_index=True
                )
            else:
                st.error("Could not connect to Google Sheets")
                
        except Exception as e:
            st.error(f"Error loading audit log: {str(e)}")

    # Footer with copyright notice
    st.markdown(
        """
        <div class="footer">
            © 2024 All rights reserved to Asaf and Matan
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
