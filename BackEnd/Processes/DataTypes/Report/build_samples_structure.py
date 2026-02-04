from collections import defaultdict

def build_samples_structure(rows):
    samples = {}

    for r in rows:
        key = r.ClientSampleID

        if key not in samples:
            samples[key] = {
                "clientSampleId": r.ClientSampleID,
                "labSampleId": r.LabSampleID,
                "dateCollected": r.DateCollected,
                "collectedBy": r.Sampler,
                "matrixId": r.MatrixID,
                "sampleTests": []
            }

        samples[key]["sampleTests"].append({
            "analyteName": r.AnalyteName,
            "results": f"{r.Result}{r.LabQualifiers}" if (r.LabQualifiers and r.LabQualifiers.strip()) else r.Result,
            "units": r.ResultUnits,
            "df": r.Dilution,
            "mdl": r.DetectionLimit,
            "pql": r.ReportingLimit,
            "method": r.LabAnalysisRefMethodID,
            "by": r.Analyst,
            "batch": r.MethodBatchID,
            "note": r.Notes,
            "groupLongName": r.GroupLongName
        })

    return list(samples.values())