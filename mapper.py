def map_to_presales(data):

    def find(section_name):
        for key in data:
            if section_name.lower() in key.lower():
                return data[key]
        return ""

    mapped = {
        "PRODUCT_SUMMARY": find("Introduction"),
        "VALUE_PROP": find("Added Value"),
        "DESCRIPTION": find("Standard Services"),
        "REQUIREMENTS": find("Prerequisites"),
        "SCOPE": find("Out of scope"),
        "SLA": find("Service Level Agreement"),
        "OPS_SUPPORT": find("Operational"),
    }

    return mapped