from typing import Sequence, Optional, cast
import transport
import mdls  # Ensure this module exists and is correctly imported

# GET /connections -> Sequence[mdls.DBConnection]
def all_connections( self, fields: Optional[str] = None,
    transport_options: Optional[transport.TransportOptions] = None,
) -> Sequence[mdls.DBConnection]:
    """Get All Connections"""
    response = cast(
        Sequence[mdls.DBConnection],
        self.get(
            path="/connections",
            structure=Sequence[mdls.DBConnection],
            query_params={"fields": fields},
            transport_options=transport_options
        )
    )
    return response