import sys
import os
from generator import ResultsGenerator
from data_loader import DataLoader
from config import Config


def main():
    print("=" * 60)
    print(" PIPELINE RESULTS GENERATOR ")
    print("=" * 60)

    # 1. Project Setup
    project_name = input("Enter Project Name (creates/uses folder in 'projects/'): ").strip()
    if not project_name:
        print("[ERROR] Project name is required.")
        return

    # Update Config paths dynamically
    Config.set_project(project_name)

    # 2. Initialize Loader
    # (It now reads files from projects/{name}/input)
    loader = DataLoader()

    try:
        # 3. Load Data
        data = loader.load_data()

        if not data:
            print("[WARNING] No data found or Group processing failed.")
            return

        # 4. Initialize Generator
        # Pass project name explicitly to handle filename generation
        generator = ResultsGenerator(project_name=project_name)

        # 5. Run Generation
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
