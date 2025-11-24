import json
import sys

def load_input(input_path):
    with open(input_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_output(output_path, decision):
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(decision, f, ensure_ascii=False, indent=4)

def make_decision(data):
    # Simplified EDE logic (from PDF: NO_ACTION, EXECUTE, POSTPONE)
    maintenance_cost = data.get('maintenance_cost', 0)
    financial_loss_without = data.get('financial_loss_without', 0)
    if financial_loss_without > maintenance_cost:
        return {"decision": "EXECUTE", "savings": financial_loss_without - maintenance_cost}
    elif financial_loss_without > 0.5 * maintenance_cost:
        return {"decision": "POSTPONE", "savings": financial_loss_without * 0.5}
    else:
        return {"decision": "NO_ACTION", "savings": 0}

if __name__ == "__main__":
    input_path = sys.argv[1] if len(sys.argv) > 1 else 'input.json'
    output_path = sys.argv[2] if len(sys.argv) > 2 else 'output.json'
    data = load_input(input_path)
    decision = make_decision(data)
    save_output(output_path, decision)
    print("Decision made:", decision)