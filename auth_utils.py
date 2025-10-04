import streamlit as st
import hashlib
import hmac
import json
import os
from datetime import datetime, timedelta
import secrets

# Configuration
AUTH_CONFIG = {
    'session_timeout_minutes': 10,
    'max_login_attempts': 3,
    'lockout_duration_minutes': 15,
    'session_refresh_threshold_minutes': 3,  # Refresh session if less than 3 minutes left
    'auto_logout_warning_minutes': 2  # Show warning when 2 minutes left
}

# Default users (in production, this should be in a secure database)
DEFAULT_USERS = {
    'aman': {
        'password_hash': 'ae2e12db5acb6575ce9bac996fe00676f680d2c773648b87485b39ba13fa7adb',  # '0506334625'
        'role': 'admin',
        'created_at': datetime.now().isoformat()
    }
}

def hash_password(password: str) -> str:
    """Hash a password using SHA-256"""
    return hashlib.sha256(password.encode()).hexdigest()

def verify_password(password: str, password_hash: str) -> bool:
    """Verify a password against its hash"""
    return hmac.compare_digest(hash_password(password), password_hash)

def load_users():
    """Load users from file or return default users"""
    users_file = 'users.json'
    if os.path.exists(users_file):
        try:
            with open(users_file, 'r') as f:
                return json.load(f)
        except:
            pass
    return DEFAULT_USERS.copy()

def save_users(users):
    """Save users to file"""
    try:
        with open('users.json', 'w') as f:
            json.dump(users, f, indent=2)
    except Exception as e:
        st.error(f"Error saving users: {e}")

def initialize_session_state():
    """Initialize authentication-related session state"""
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'username' not in st.session_state:
        st.session_state.username = None
    if 'user_role' not in st.session_state:
        st.session_state.user_role = None
    if 'login_time' not in st.session_state:
        st.session_state.login_time = None
    if 'login_attempts' not in st.session_state:
        st.session_state.login_attempts = {}
    if 'session_token' not in st.session_state:
        st.session_state.session_token = None

def is_session_valid():
    """Check if the current session is valid"""
    if not st.session_state.authenticated:
        return False
    
    if not st.session_state.login_time:
        return False
    
    # Check session timeout
    login_time = datetime.fromisoformat(st.session_state.login_time)
    timeout_duration = timedelta(minutes=AUTH_CONFIG['session_timeout_minutes'])
    session_duration = datetime.now() - login_time
    
    if session_duration > timeout_duration:
        logout_user()
        st.warning("⚠️ Session expired. Please login again.")
        return False
    
    # Show warning if session is about to expire
    minutes_left = AUTH_CONFIG['session_timeout_minutes'] - int(session_duration.total_seconds() / 60)
    if minutes_left <= AUTH_CONFIG['auto_logout_warning_minutes']:
        st.warning(f"⚠️ Session will expire in {minutes_left} minutes. Please save your work.")
    
    return True

def refresh_session():
    """Refresh the current session"""
    if st.session_state.authenticated:
        st.session_state.login_time = datetime.now().isoformat()
        return True
    return False

def get_session_info():
    """Get current session information"""
    if not st.session_state.authenticated:
        return None
    
    login_time = datetime.fromisoformat(st.session_state.login_time)
    session_duration = datetime.now() - login_time
    minutes_left = AUTH_CONFIG['session_timeout_minutes'] - int(session_duration.total_seconds() / 60)
    
    return {
        'username': st.session_state.username,
        'role': st.session_state.user_role,
        'login_time': login_time,
        'session_duration': session_duration,
        'minutes_left': max(0, minutes_left),
        'needs_refresh': minutes_left <= AUTH_CONFIG['session_refresh_threshold_minutes']
    }

def is_user_locked_out(username: str) -> bool:
    """Check if user is locked out due to failed login attempts"""
    if username not in st.session_state.login_attempts:
        return False
    
    attempts_data = st.session_state.login_attempts[username]
    if attempts_data['count'] < AUTH_CONFIG['max_login_attempts']:
        return False
    
    # Check if lockout period has expired
    lockout_time = datetime.fromisoformat(attempts_data['lockout_time'])
    lockout_duration = timedelta(minutes=AUTH_CONFIG['lockout_duration_minutes'])
    
    if datetime.now() - lockout_time > lockout_duration:
        # Reset attempts after lockout period
        st.session_state.login_attempts[username] = {'count': 0, 'lockout_time': None}
        return False
    
    return True

def record_failed_login(username: str):
    """Record a failed login attempt"""
    if username not in st.session_state.login_attempts:
        st.session_state.login_attempts[username] = {'count': 0, 'lockout_time': None}
    
    st.session_state.login_attempts[username]['count'] += 1
    
    if st.session_state.login_attempts[username]['count'] >= AUTH_CONFIG['max_login_attempts']:
        st.session_state.login_attempts[username]['lockout_time'] = datetime.now().isoformat()

def reset_login_attempts(username: str):
    """Reset login attempts for successful login"""
    if username in st.session_state.login_attempts:
        st.session_state.login_attempts[username] = {'count': 0, 'lockout_time': None}

def authenticate_user(username: str, password: str) -> bool:
    """Authenticate a user with username and password"""
    if is_user_locked_out(username):
        st.error(f"Account locked due to multiple failed attempts. Try again in {AUTH_CONFIG['lockout_duration_minutes']} minutes.")
        return False
    
    users = load_users()
    
    if username not in users:
        record_failed_login(username)
        st.error("Invalid username or password")
        return False
    
    if not verify_password(password, users[username]['password_hash']):
        record_failed_login(username)
        st.error("Invalid username or password")
        return False
    
    # Successful login
    reset_login_attempts(username)
    st.session_state.authenticated = True
    st.session_state.username = username
    st.session_state.user_role = users[username]['role']
    st.session_state.login_time = datetime.now().isoformat()
    st.session_state.session_token = secrets.token_urlsafe(32)
    
    return True

def logout_user():
    """Logout the current user"""
    st.session_state.authenticated = False
    st.session_state.username = None
    st.session_state.user_role = None
    st.session_state.login_time = None
    st.session_state.session_token = None

def require_authentication():
    """Decorator function to require authentication for app sections"""
    if not is_session_valid():
        return False
    return True

def show_login_form():
    """Display the login form"""
    st.markdown("## 🔐 Login Required")
    st.markdown("Please login to access the Supplier Quote Optimizer")
    
    with st.form("login_form"):
        col1, col2, col3 = st.columns([1, 2, 1])
        
        with col2:
            st.markdown("### Login")
            username = st.text_input("Username", placeholder="Enter your username")
            password = st.text_input("Password", type="password", placeholder="Enter your password")
            
            submitted = st.form_submit_button("Login", use_container_width=True)
            
            if submitted:
                if not username or not password:
                    st.error("Please enter both username and password")
                else:
                    if authenticate_user(username, password):
                        st.success("Login successful!")
                        st.rerun()

def show_user_info():
    """Display current user information and logout option"""
    if st.session_state.authenticated:
        col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
        
        with col1:
            st.markdown(f"👤 **Logged in as:** {st.session_state.username} ({st.session_state.user_role})")
        
        with col2:
            session_info = get_session_info()
            if session_info:
                st.markdown(f"⏱️ **Session:** {session_info['minutes_left']}min left")
                if session_info['needs_refresh']:
                    if st.button("🔄 Refresh Session", type="secondary", help="Extend your session"):
                        refresh_session()
                        st.success("✅ Session refreshed!")
                        st.rerun()
        
        with col3:
            # Admin Panel button (only for admin users)
            if st.session_state.user_role == 'admin':
                if st.button("👥 Admin Panel", type="secondary"):
                    st.session_state.show_admin_panel = not st.session_state.get('show_admin_panel', False)
                    st.rerun()
        
        with col4:
            if st.button("Logout", type="secondary"):
                logout_user()
                st.rerun()
        
        # Admin Panel (only for admin users)
        if st.session_state.get('show_admin_panel', False) and st.session_state.user_role == 'admin':
            st.markdown("---")
            st.markdown("### 👥 User Management Panel")
            
            # Add new user
            with st.expander("➕ Add New User", expanded=False):
                with st.form("add_user_form"):
                    new_username = st.text_input("Username", placeholder="Enter username")
                    new_password = st.text_input("Password", type="password", placeholder="Enter password")
                    confirm_password = st.text_input("Confirm Password", type="password", placeholder="Confirm password")
                    new_role = st.selectbox("Role", ["user", "admin"], index=0)
                    
                    if st.form_submit_button("Add User"):
                        if new_username and new_password:
                            if new_password == confirm_password:
                                if create_new_user(new_username, new_password, new_role):
                                    st.success(f"✅ User '{new_username}' added successfully!")
                                    st.rerun()
                                else:
                                    st.error("❌ Username already exists!")
                            else:
                                st.error("❌ Passwords do not match!")
                        else:
                            st.error("❌ Please fill in all fields!")
            
            # List existing users
            st.markdown("#### 📋 Existing Users")
            users_data = get_all_users()
            if users_data:
                for username, user_info in users_data.items():
                    col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
                    with col1:
                        role_icon = "👑" if user_info.get('role', 'user') == 'admin' else "👤"
                        st.write(f"{role_icon} {username} ({user_info.get('role', 'user')})")
                    
                    with col2:
                        if username != 'admin':  # Don't allow deleting admin
                            if st.button(f"🗑️ Delete", key=f"delete_{username}"):
                                if delete_user(username):
                                    st.success(f"✅ User '{username}' deleted!")
                                    st.rerun()
                                else:
                                    st.error(f"❌ Failed to delete user '{username}'!")
                    
                    with col3:
                        if st.button(f"🔄 Reset Password", key=f"reset_{username}"):
                            st.session_state[f'reset_password_{username}'] = True
                            st.rerun()
                    
                    with col4:
                        if username != 'admin' and user_info.get('role', 'user') != 'admin':
                            if st.button(f"👑 Make Admin", key=f"promote_{username}"):
                                if update_user_role(username, 'admin'):
                                    st.success(f"✅ User '{username}' promoted to admin!")
                                    st.rerun()
                                else:
                                    st.error(f"❌ Failed to promote user '{username}'!")
                    
                    # Password reset form
                    if st.session_state.get(f'reset_password_{username}', False):
                        with st.form(f"reset_password_form_{username}"):
                            new_pwd = st.text_input("New Password", type="password", key=f"new_pwd_{username}")
                            confirm_pwd = st.text_input("Confirm Password", type="password", key=f"confirm_pwd_{username}")
                            
                            col_submit, col_cancel = st.columns(2)
                            with col_submit:
                                if st.form_submit_button("Update Password"):
                                    if new_pwd and new_pwd == confirm_pwd:
                                        if update_user_password(username, new_pwd):
                                            st.success(f"✅ Password updated for '{username}'!")
                                            st.session_state[f'reset_password_{username}'] = False
                                            st.rerun()
                                        else:
                                            st.error("❌ Failed to update password!")
                                    else:
                                        st.error("❌ Passwords do not match or are empty!")
                            with col_cancel:
                                if st.form_submit_button("Cancel"):
                                    st.session_state[f'reset_password_{username}'] = False
                                    st.rerun()
            else:
                st.info("No users found.")
            
            st.markdown("---")

def create_new_user(username: str, password: str, role: str = 'user') -> bool:
    """Create a new user (admin only)"""
    if st.session_state.user_role != 'admin':
        st.error("Only administrators can create new users")
        return False
    
    users = load_users()
    
    if username in users:
        return False
    
    users[username] = {
        'password_hash': hash_password(password),
        'role': role,
        'created_at': datetime.now().isoformat()
    }
    
    save_users(users)
    return True

def get_all_users():
    """Get all users (admin only)"""
    if st.session_state.user_role != 'admin':
        return {}
    return load_users()

def delete_user(username: str) -> bool:
    """Delete a user (admin only)"""
    if st.session_state.user_role != 'admin':
        return False
    
    if username == 'admin':  # Protect admin account
        return False
    
    users = load_users()
    if username in users:
        del users[username]
        save_users(users)
        return True
    return False

def update_user_password(username: str, new_password: str) -> bool:
    """Update user password (admin only)"""
    if st.session_state.user_role != 'admin':
        return False
    
    users = load_users()
    if username in users:
        users[username]['password_hash'] = hash_password(new_password)
        save_users(users)
        return True
    return False

def update_user_role(username: str, new_role: str) -> bool:
    """Update user role (admin only)"""
    if st.session_state.user_role != 'admin':
        return False
    
    if username == 'admin':  # Protect admin account
        return False
    
    users = load_users()
    if username in users:
        users[username]['role'] = new_role
        save_users(users)
        return True
    return False

def show_user_management():
    """Show user management interface (admin only)"""
    if st.session_state.user_role != 'admin':
        return
    
    st.markdown("### 👥 User Management")
    
    users = load_users()
    
    # Display existing users
    st.markdown("**Existing Users:**")
    for username, user_data in users.items():
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.text(f"{username} ({user_data['role']})")
        with col2:
            created = datetime.fromisoformat(user_data['created_at']).strftime("%Y-%m-%d")
            st.text(f"Created: {created}")
    
    # Add new user form
    with st.expander("Add New User"):
        with st.form("new_user_form"):
            new_username = st.text_input("New Username")
            new_password = st.text_input("New Password", type="password")
            new_role = st.selectbox("Role", ["user", "admin"])
            
            if st.form_submit_button("Create User"):
                if new_username and new_password:
                    if create_new_user(new_username, new_password, new_role):
                        st.success(f"User '{new_username}' created successfully!")
                        st.rerun()
                else:
                    st.error("Please enter both username and password")