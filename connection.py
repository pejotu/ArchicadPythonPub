"""ArchiCAD connection management with Tapir addon validation."""

from archicad import ACConnection


class ArchicadConnection:
    """Wraps ACConnection and validates that the Tapir addon is available."""

    def __init__(self):
        self.conn = None
        self.cmd = None
        self.types = None

    def connect(self):
        """Establish connection to the running ArchiCAD instance."""
        self.conn = ACConnection.connect()
        if not self.conn:
            raise ConnectionError(
                "Could not connect to ArchiCAD. "
                "Make sure ArchiCAD is running and the Python API is enabled."
            )
        self.cmd = self.conn.commands
        self.types = self.conn.types
        self._check_tapir()
        return self

    def _check_tapir(self):
        """Verify the Tapir addon is loaded and responding."""
        try:
            self.cmd.ExecuteAddOnCommand(
                self.types.AddOnCommandId("TapirCommand", "GetAddOnVersion")
            )
        except Exception as exc:
            raise RuntimeError(
                "Tapir addon is not available. "
                "Install it from https://github.com/ENZYME-APD/tapir-archicad-automation"
            ) from exc

    def tapir(self, command: str, parameters=None):
        """Execute a Tapir addon command.

        Args:
            command:    Tapir command name (e.g. 'GetGeoLocation').
            parameters: Optional dict of command parameters.

        Returns:
            Command result dict.
        """
        cmd_id = self.types.AddOnCommandId("TapirCommand", command)
        if parameters is not None:
            return self.cmd.ExecuteAddOnCommand(cmd_id, parameters)
        return self.cmd.ExecuteAddOnCommand(cmd_id)
