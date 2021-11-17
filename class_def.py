


class Part(object):
    """Object to represent a part, assy, or mod.
    """
    def __init__(self, part_num, Parents=set()):
        self.part_num = part_num
        self.Parents = Parents

    def get_obs_status(self):
        self.can_obs = True
        for Parent in self.Parents:
            parent_status = Parent.get_obs_status()
            if parent_status == False:
                self.can_obs = False
                break
        # If no parents in set, results in self.can_obs = True as it should.
        return self.can_obs

    def add_parent(self, Parent):
        self.Parents.add(Parent)

    def get_parents(self):
        return self.Parents

    def get_pn(self):
        return self.part_num

    def __str__(self):
        return self.part_num

    def __repr__(self):
        return "Part object: %s" % self.part_num


class Platform(Part):
    """Object to represent a platform. Inherits from Part class.
    """
    def __init__(self, part_num, can_obs):
        self.part_num = part_num
        self.Parents = None
        self.can_obs = can_obs

    def get_obs_status(self):
        return self.can_obs

    def __repr__(self):
        return "Platform object: %s" % self.part_num

# testing
# Pltfm1 = Platform("658237", True)
# # print(Pltfm1, ": ", Pltfm1.get_obs_status())
# Pltfm2 = Platform("658244", False)
# Pltfm3 = Platform("10016534", True)
#
# Part3 = Part("10013547", Parents=set({Pltfm1, Pltfm3}))
# print(Part3, ": ", Part3.get_obs_status())
#
# Part4 = Part("654981", Parents=set({Pltfm2}))
# print(Part4, ": ", Part4.get_obs_status())
#
# Part3.add_parent(Part4)
# print(Part3, ": ", Part3.get_obs_status())
