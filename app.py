def calculate_confidence_score(data):
    required_fields = ["international depositories", "local agent bank", "securities a/c no"]
    optional_fields = ["currency", "beneficiary_name"]

    score = 0
    missing_fields = []

    for field in required_fields:
        if data.get(field):
            score += 2
        else:
            missing_fields.append(field)

    for field in optional_fields:
        if data.get(field):
            score += 1
        else:
            missing_fields.append(field)

    max_score = 2 * len(required_fields) + len(optional_fields)
    confidence = score / max_score if max_score > 0 else 0

    return {
        "document_is_ssi": confidence >= 0.6,  # You can adjust this threshold
        "confidence_score": round(confidence, 2),
        "missing_fields": missing_fields
    }
