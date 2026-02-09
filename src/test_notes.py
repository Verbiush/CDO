
import json

def recursive_update_notes(obj, tipo_val, num_val):
    count = 0
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k == "tipoNota" and tipo_val:
                obj[k] = tipo_val
                count += 1
            elif k == "numNota" and num_val:
                obj[k] = num_val
                count += 1
            else:
                count += recursive_update_notes(v, tipo_val, num_val)
    elif isinstance(obj, list):
        for v in obj:
            count += recursive_update_notes(v, tipo_val, num_val)
    return count

def test_logic():
    data = {
        "usuarios": [
            {
                "nombre": "Juan",
                "tipoNota": None,
                "numNota": None,
                "servicios": [
                    {
                        "codServicio": "123",
                        "tipoNota": "OLD",
                        "numNota": "OLD-01"
                    }
                ]
            }
        ],
        "meta": {
            "tipoNota": "ROOT",
            "numNota": "ROOT-01"
        }
    }

    print("Original:", json.dumps(data, indent=2))
    
    # Update
    changes = recursive_update_notes(data, "NA", "NA-01")
    print(f"Changes: {changes}")
    print("Updated:", json.dumps(data, indent=2))
    
    # Assertions
    assert data["usuarios"][0]["tipoNota"] == "NA"
    assert data["usuarios"][0]["numNota"] == "NA-01"
    assert data["usuarios"][0]["servicios"][0]["tipoNota"] == "NA"
    assert data["usuarios"][0]["servicios"][0]["numNota"] == "NA-01"
    assert data["meta"]["tipoNota"] == "NA"
    assert data["meta"]["numNota"] == "NA-01"
    print("Test Passed!")

if __name__ == "__main__":
    test_logic()
