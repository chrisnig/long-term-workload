from util.filter import TupleFilter


class DataSet:
    """Represents a set of several variables.
    """

    TYPE_MATRIX = "matrix"
    TYPE_SET = "set"
    TYPE_SCALAR = "scalar"

    def __init__(self) -> None:
        self.variables = dict()
        self.matching_cache = dict()

    def create(self, var: str, default=0, dimensions=()) -> None:
        """Creates a new variable in the solution with all indices initialized to the given default value.
        """

        self._checkVarDoesNotExist(var)

        self.variables[var] = {
            "default": default,
            "type": self.TYPE_MATRIX,
            "values": dict(),
            "filter": None,
            "dimensions": dimensions
        }

    def create_set(self, var: str, value=set()) -> None:
        """Creates a new set variable in the data initialized with the supplied set.
        """

        self._checkVarDoesNotExist(var)

        self.variables[var] = {
            "type": self.TYPE_SET,
            "values": value
        }

    def create_scalar(self, var: str, value=0) -> None:
        """Creates a new scalar variable in the data
        """

        self._checkVarDoesNotExist(var)

        self.variables[var] = {
            "type": self.TYPE_SCALAR,
            "values": value
        }

    def remove(self, var: str) -> None:
        """Removes a variable from the data set
        """

        del self.variables[var]

    def get_set(self, var: str) -> set:
        """Gets the value of a set variable.
        """

        return self.variables[var]["values"]

    def get_scalar(self, var: str) -> float:
        """Gets the value of a scalar variable
        """

        return self.variables[var]["values"]

    def get(self, var: str, index) -> any:
        """Gets the value of a variable at a given index.
        """

        if index in self.variables[var]["values"]:
            return self.variables[var]["values"][index]

        return self.variables[var]["default"]

    def get_dimensions(self, var: str) -> tuple:
        """Gets the dimensions of a variable.
        """

        return self.variables[var]["dimensions"]

    def get_values(self, var: str) -> dict:
        """Gets all entries of a variable which are not equal to the variable's default value/
        """

        return self.variables[var]["values"]

    def get_matching_keys(self, var: str, match: tuple) -> set:
        """Checks all non-default entries of a variable using the given filter. Matching is done for each index.
        Example: Filter (1, None, 2) will match all entries which have the value 1 at the first index, an arbitrary
        value at the second index and the value 2 at the third index.
        
        Note that the dimensions of the filter must match the dimensions of the variable.
        """

        if self.variables[var]["filter"] is None:
            self.variables[var]["filter"] = TupleFilter(self.variables[var]["values"])

        return self.variables[var]["filter"].arity_filtered(match)

    def set(self, var: str, index, value: any) -> None:
        """Sets the value of a variable at a given index.
        """

        if value == self.variables[var]["default"]:
            if index in self.variables[var]["values"]:
                del self.variables[var]["values"][index]

                if self.variables[var]["filter"] is not None:
                    self.variables[var]["filter"].remove(index)
        else:
            self.variables[var]["values"][index] = value

            if self.variables[var]["filter"] is not None:
                self.variables[var]["filter"].add(index)

    def set_values(self, var: str, values: dict) -> None:
        """Sets all the values of the variable to the supplied dictionary. Note that this overrides _all_ existing
        values.
        """

        self.variables[var]["values"] = values

    def __contains__(self, var: str) -> bool:
        """Checks whether a variable exists in this solution
        """

        return var in self.variables

    def get_variables_by_type(self, variable_type: int) -> iter:
        """Gets all variables of a given type. Type must be one of the variable types specified in constants at the top
        of this class.
        """
        return ((name, value) for name, value in self.variables.items() if value["type"] == variable_type)

    def _checkVarDoesNotExist(self, var: str) -> None:
        if var in self:
            raise RuntimeError("Trying to create variable {} which already exists.".format(var))

    def count_if_exists(self, var: str, match: tuple) -> int:
        """"Gets the number of keys for which a value is set on a variable. Returns 0 if the variable does not exist.
        """
        if var in self:
            return len(self.get_matching_keys(var, match))
        else:
            return 0
