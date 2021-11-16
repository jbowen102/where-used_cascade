


class Part(object):
    """Object to represent a part, assy, or mod.
    """
    def __init__(self, part_num, Parents=set()):
        # try to get obs status.
        self.part_num = part_num
        self.Parents = Parents
        # self.Children = set()
        # for Parent in self.Parents:
        #     Parent.add_child(self)
        self.update_obs_status() # defines self.can_obs

    def update_obs_status(self):
        self.can_obs = True
        for Parent in self.Parents:
            parent_status = Parent.get_obs_status()
            if parent_status == False:
                self.can_obs = False
                break
        # If no parents in set, results in self.can_obs = True as it should.
        # self.update_children()

    # def update_children(self):
    #     for Child in self.Children:
    #         Child.update_obs_status()

    def get_obs_status(self):
        # if self.can_obs == None:
        #     self.update_obs_status()
        # return self.can_obs
        self.update_obs_status()
        return self.can_obs

    def add_parent(self, Parent):
        self.Parents.add(Parent)
        self.update_obs_status() # update this part's status.

    # def add_child(self, Child):
    #     self.Children.add(Child)

    def get_parents(self):
        return self.Parents

    # def get_children(self):
    #     return self.Children

    def get_pn(self):
        return self.part_num

    def __str__(self):
        return self.part_num

    def __repr__(self):
        return "Part object: %s" % self.part_num


class Platform(Part):
    """Object to represent a platform. Inherits from Part class.
    """
    def __init__(self, part_num, obs_status):
        self.part_num = part_num
        self.Parents = None
        self.can_obs = obs_status

    def update_obs_status(self):
        pass
