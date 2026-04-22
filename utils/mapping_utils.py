import json
 
def load_mapping():

    with open("config/mapping.json") as f:

        return json.load(f)
 
def reverse_mapping(mapping):

    return {v: k for k, v in mapping.items()}