

class Result:
    """
    Class containing all the relevant extracted and calculated data for a single subject.
    """
    def __init__(self, mdate, rat, dur, var_per, tot_al, tot_il, lat_r1, lat_fr10a,
                 lat_fr10i, lat_fr10a_aft_vi, lat_fr10i_aft_vi, vi, num_vi_al, arr_al,
                 num_vi_il, arr_il, rew, mag, al_vi_to_rew, il_vi_to_rew):
        self.date = mdate
        self.rat = rat
        self.dur = dur
        self.vp = var_per
        self.tot_al = tot_al
        self.tot_il = tot_il
        self.lat_r1 = lat_r1
        self.lat_fr10a = lat_fr10a
        self.lat_fr10i = lat_fr10i
        self.lat_fr10a_aft_vi = lat_fr10a_aft_vi
        self.lat_fr10i_aft_vi = lat_fr10i_aft_vi
        self.vi = vi
        self.num_vi_al = num_vi_al
        self.arr_al = arr_al
        self.num_vi_il = num_vi_il
        self.arr_il = arr_il
        self.rew = rew
        self.mag = mag
        self.al_vi_to_rew = al_vi_to_rew
        self.il_vi_to_rew = il_vi_to_rew
