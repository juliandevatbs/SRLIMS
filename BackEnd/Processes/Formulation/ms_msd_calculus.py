def ms_msd_calculus(result, parent_result, parent_qualifiers, parent_detection_limit, qc_spike):
      
    if qc_spike is None or qc_spike == 0:
        raise ValueError("QCSpikeAdded cannot be 0 or empty")
    
    try:
        result_float = float(result)
        qc_spike_float = float(qc_spike)
    except (ValueError, TypeError) as e:
        raise ValueError(f"Invalid numeric values: {e}")
    
    if parent_qualifiers == 'U':
        if parent_detection_limit is None:
            raise ValueError("Parent sample has 'U' qualifier but no Detection Limit")
        try:
            parent_value = float(parent_detection_limit)
        except (ValueError, TypeError):
            raise ValueError("Invalid Detection Limit value")
    else:
        try:
            parent_value = float(parent_result)
        except (ValueError, TypeError):
            raise ValueError("Invalid parent result value")
    
    #  % Recovery: ((MS/MSD - Parent) / Spike) * 100
    percent_recovery = ((result_float - parent_value) / qc_spike_float) * 100
    
    # Redondear a 2 decimales
    return round(percent_recovery, 2)