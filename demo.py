"""
demo.py -- runs the valuation engine, then the deal room generator that
reads its output, against the illustrative sample_data/ tax package (see
build_sample_data.py). No real business data required.

    pip install -r requirements.txt
    python build_sample_data.py
    python demo.py
"""
from mna_valuation_engine import calculate_mna_valuation
from mna_deal_room_generator import generate_deal_room_package

if __name__ == "__main__":
    print("### Step 1 -- Valuation Engine ###\n")
    calculate_mna_valuation()
    print("\n### Step 2 -- Deal Room Generator ###\n")
    generate_deal_room_package()

