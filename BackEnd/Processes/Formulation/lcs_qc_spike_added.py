def lcs_qc_spike_added(result, qc_spike):
    """
    QCSpikeAdded (%) = (Result / QCSpike) * 100
    """
    try:
        result_val = float(result)
        spike_val = float(qc_spike)

        if spike_val == 0:
            return None

        return (result_val / spike_val) * 100

    except (TypeError, ValueError) as e:
        
        print(f"Error calculating QCSpikeAdded: {e}")
        
        return None
