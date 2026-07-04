import pandas as pd

from app import build_competency_matrix_snapshot


def test_build_competency_matrix_snapshot_uses_current_phase4_tables():
    training_records = pd.DataFrame(
        [{"user_id": "U1", "test_status": "Passed", "status": "Completed"}]
    )
    witness_surveys = pd.DataFrame(
        [{"user_id": "U1", "scope": "In-Service Surveyor - Hull", "outcome": "Pass"}]
    )
    supervised_activities = pd.DataFrame(
        [{"user_id": "U1", "scope": "In-Service Surveyor - Hull", "activity_kind": "Independent Survey", "outcome": "Pass"}]
    )
    authorization_requests = pd.DataFrame(
        [{"user_id": "U1", "scope": "In-Service Surveyor - Hull", "status": "Management Approved"}]
    )

    result = build_competency_matrix_snapshot(
        user_id="U1",
        scope="In-Service Surveyor - Hull",
        pathway="In-Service Surveyor",
        training_records=training_records,
        witness_records=witness_surveys,
        supervised_records=supervised_activities,
        authorization_requests=authorization_requests,
        development_plans=pd.DataFrame(),
    )

    assert result["training_status"] == "Completed"
    assert result["mcq_status"] == "Passed"
    assert result["witness_completed"] == 1
    assert result["supervised_completed"] == 1
    assert result["authorization_status"] == "Authorized"
    assert result["risk_level"] == "Low"
    assert result["gap_summary"] == "No major gap"
