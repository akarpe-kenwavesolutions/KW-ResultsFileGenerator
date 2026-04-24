import sys
import os
from generator import ResultsGenerator
from data_loader import DataLoader
from config import Config

def main():
    print("=" * 60)
    print("         KW RESULTS FILE TEMPLATE GENERATOR          ")
    print("=" * 60)

    # 1. Project Setup
    project_name = input("Enter Project Name (creates/uses folder in 'projects/'): ").strip()
    if not project_name:
        print("[ERROR] Project name is required.")
        return

    # Update Config paths dynamically
    Config.set_project(project_name)

    # 2. Ask about Pipe Asset IDs BEFORE loading data
    print("\n--- Pipe Asset IDs Configuration ---")
    while True:
        asset_response = input("Include Pipe Asset IDs in this project? [y/n]: ").strip().lower()
        if asset_response in ['y', 'yes']:
            Config.REQUIRE_ASSET_IDS = True
            print(">> Pipe Asset IDs will be included.")
            break
        elif asset_response in ['n', 'no']:
            Config.REQUIRE_ASSET_IDS = False
            print(">> Pipe Asset IDs will be excluded.")
            break
        else:
            print("Invalid input. Please enter 'y' or 'n'.")

    # 3. Initialize Loader (now it knows whether to require asset file)
    loader = DataLoader()

    try:
        # 4. Load Data
        data = loader.load_data()
        if not data:
            print("[WARNING] No data found or Group processing failed.")
            return

        # 5. Initialize Generator
        generator = ResultsGenerator(project_name=project_name)

        # 6. Run Generation
        generator.run(data)

    except FileNotFoundError as e:
        print(f"\n[ERROR] File not found: {e}")
    except ValueError as e:
        print(f"\n[ERROR] Data validation failed: {e}")
    except Exception as e:
        print(f"\n[CRITICAL ERROR] {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
