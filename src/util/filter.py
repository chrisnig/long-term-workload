# TupleFilter class inspired by Nick (https://stackoverflow.com/users/2480947)
# https://stackoverflow.com/a/44321960
# Permission was given by the original author to use this code under the MIT license.

from collections import defaultdict


class TupleFilter:
    def __init__(self, data: dict) -> None:
        self.data = data
        self.lookup = self._build_lookup()

    def _build_lookup(self) -> dict:
        lookup = defaultdict(set)
        for data_item in self.data:
            for member_ref, data_key in TupleFilter._tuple_index(data_item).items():
                lookup[member_ref].add(data_key)
        return lookup

    @staticmethod
    def _tuple_index(item_key: tuple) -> dict:
        member_refs = enumerate(item_key)
        return {(pos, val): item_key for pos, val in member_refs}

    def filtered(self, tuple_filter: tuple) -> set:
        # initially unfiltered
        results = self.all_keys()
        # reduce filtered set
        for position, value in enumerate(tuple_filter):
            if value is not None:
                match_or_empty_set = self.lookup.get((position, value), set())
                results = results.intersection(match_or_empty_set)
        return results

    def arity_filtered(self, tuple_filter: tuple) -> set:
        tf_length = len(tuple_filter)
        return {match for match in self.filtered(tuple_filter) if tf_length == len(match)}

    def all_keys(self) -> set:
        return set(self.data.keys())

    def add(self, key: tuple) -> None:
        for member_ref, data_key in TupleFilter._tuple_index(key).items():
            self.lookup[member_ref].add(key)

    def remove(self, key: tuple) -> None:
        for member_ref, data_key in TupleFilter._tuple_index(key).items():
            self.lookup[member_ref].remove(key)
