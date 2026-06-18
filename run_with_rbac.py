import os
import sys

# Ensure the app_code directory is in path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from loginpage import LoginPage
from mainapp import MainApp
from rbac_manager import RBACManager

def launch_app_with_rbac():
    """
    Launches the application with Role-Based Access Control
    without modifying the original source files.
    """
    print("Initializing RBAC System...")
    rbac_manager = RBACManager()
    
    # We will monkey-patch LoginPage to use the RBAC authentication
    # and hook into the open_main_app to apply UI restrictions.
    
    original_login = LoginPage.login
    original_open_main_app = LoginPage.open_main_app
    
    def patched_login(self):
        # Retrieve credentials from the entry fields (using the RoundedEntry's get method)
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if not username or not password:
            self.error_label.config(text="Please enter username and password")
            return
            
        # Use our RBACManager to authenticate, which merges custom_users and Config.USERS
        success, result = rbac_manager.authenticate(username, password)
        
        if not success:
            self.error_label.config(text=result)
            return
            
        self.logger.info(f"User logged in via RBAC: {username}")
        # Store the authenticated user to be accessed in the next step
        self._authenticated_username = username
        self.open_main_app(result["company"])
        
    def patched_open_main_app(self, company):
        # Call the original to spawn MainApp
        original_open_main_app(self, company)
        
        # Now apply RBAC restrictions to the freshly created MainApp
        username = getattr(self, '_authenticated_username', None)
        if username and self.current_app:
            print(f"Applying RBAC rules for user: {username}")
            rbac_manager.enforce_rbac(self.current_app, username)

    # Apply the patches
    LoginPage.login = patched_login
    LoginPage.open_main_app = patched_open_main_app
    
    # Start the app
    app = LoginPage()
    app.mainloop()

if __name__ == "__main__":
    launch_app_with_rbac()
