# Copyright (c) 2026 Μαρία Ελένη Αντωνοπούλου
# Licensed under the MIT License. See LICENSE file.
#
# Unified scraper for 2026 U18 100m & 200m Greek performances.
# Data sourced from Roster Athletics (https://www.rosterathletics.com).

import sys

def select_event():
    """
    Interactive menu to select which event to process.
    Returns: '100m' or '200m'
    """
    print("\n" + "="*60)
    print("   2026 U18 Greek Track & Field - Performance Tracker")
    print("="*60)
    print("\nSelect Event:")
    print("  1. 100m")
    print("  2. 200m")
    print("  q. Quit")
    
    while True:
        choice = input("\n>> ").strip().lower()
        if choice in ('1', '100m', '100'):
            return '100m'
        elif choice in ('2', '200m', '200'):
            return '200m'
        elif choice == 'q':
            print("Goodbye!")
            sys.exit(0)
        else:
            print("Invalid choice. Please enter 1, 2, or q.")

if __name__ == "__main__":
    event = select_event()
    
    if event == '100m':
        print("\nRunning 100m scraper...")
        import scraper
    elif event == '200m':
        print("\nRunning 200m scraper...")
        import scraper_200m
