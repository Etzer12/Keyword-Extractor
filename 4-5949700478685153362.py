from strsimpy.levenshtein import Levenshtein
from strsimpy.normalized_levenshtein import *
from strsimpy.weighted_levenshtein import *
from strsimpy.damerau import Damerau
from strsimpy.jaro_winkler import JaroWinkler
from strsimpy.cosine import Cosine
from strsimpy.jaccard import Jaccard
from strsimpy.shingle_based import ShingleBased


class levenshtein:
    levenshtein_obj = Levenshtein()
    normal_levenshtein_obj = NormalizedLevenshtein()
    weighted_levenshtein_obj = WeightedLevenshtein()
    damerau_obj = Damerau()

    def find_distance(self, first_str, second_str):
        return self.levenshtein_obj.distance(first_str, second_str)

    def find_normal_distance(self, first_str, second_str):
        return self.normal_levenshtein_obj.distance(first_str, second_str)

    def find_normal_similarity(self, first_str, second_str):
        return self.normal_levenshtein_obj.similarity(first_str, second_str)

    def find_weighted(self, first_str, second_str):
        return self.weighted_levenshtein_obj.distance(first_str, second_str)

    def find_damerau(self, firs_str, second_srt):
        return self.damerau_obj.distance(firs_str, second_srt)


class jaro_winkler:
    jaro_winkler_obj = JaroWinkler()

    def find_similarity(self, first_str, second_str):
        return self.jaro_winkler_obj.similarity(first_str, second_str)

    def find_distance(self, first_str, second_str):
        return self.jaro_winkler_obj.distance(first_str, second_str)


class cosine:
    shingle_based_obj = ShingleBased()
    cosine_obj = Cosine(shingle_based_obj)

    def find_similarity(self, first_str, second_str):
        return self.cosine_obj.similarity(first_str, second_str)

    def find_distance(self, first_str, second_str):
        return self.cosine_obj.distance(first_str, second_str)

    def find_similarity_profile(self, first_str, second_str):
        return self.cosine_obj.similarity_profiles(self.cosine_obj.get_profile(first_str), self.cosine_obj.get_profile(second_str))


class jaccard_index:
    shingle_based_obj = ShingleBased()
    jaccard_index_obj = Jaccard(shingle_based_obj)

    def find_distance(self, first_str, second_str):
        return self.jaccard_index_obj.distance(first_str, second_str)

    def find_similarity(self, first_str, second_str):
        return self.jaccard_index_obj.similarity(first_str, second_str)



