import os
import sys

def ensure_trailing_slash(path):
    """
    Ensures the path ends with a backslash.
    """
    return path if path.endswith('\\') else path + '\\'

def get_app_location():
    """
    Returns a dictionary with different paths related to the application:
    - app_dir: Directory containing the script (with trailing backslash)
    - app_path: Full path to the script
    - app_name: Name of the script file
    - exe_dir: Directory containing the executable (with trailing backslash)
    """
    # Get the script path
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        app_path = sys.executable
        app_dir = ensure_trailing_slash(os.path.dirname(sys.executable))
        exe_dir = ensure_trailing_slash(os.path.dirname(sys.executable))
    else:
        # Running as script
        app_path = os.path.abspath(__file__)
        app_dir = ensure_trailing_slash(os.path.dirname(app_path))
        exe_dir = ensure_trailing_slash(os.path.dirname(sys.executable))

    return {
        'app_dir': app_dir,
        'app_path': app_path,
        'app_name': os.path.basename(app_path),
        'exe_dir': exe_dir
    }

def get_resource_path(relative_path):
    """
    Returns the absolute path to a resource file, whether running as script or frozen executable.

    Args:
        relative_path (str): Path relative to the application directory

    Returns:
        str: Absolute path to the resource
    """
    base_path = get_app_location()['app_dir']
    # base_path já terá a barra no final
    return os.path.join(base_path, relative_path)

# Example usage
if __name__ == "__main__":
    locations = get_app_location()
    print("\nApplication Locations:")
    for key, value in locations.items():
        print(f"{key}: {value}")

    # Example of getting a resource path
    resource = get_resource_path("resources/config.ini")
    print(f"\nExample resource path: {resource}")
