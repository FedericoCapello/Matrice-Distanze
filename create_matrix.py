import googlemaps
import pandas as pd
from datetime import datetime


class DM(object):
    def __init__(self):
        """Initialize a Distance Matrix Class exploiting Google Maps"""

        self.gmaps = googlemaps.Client(key='AIzaSyDhxuJ-9v2pxCyo5B0xxrKXXjYige4uIQU')
        self.distances = []
        self.now = datetime.now()
        self.true_origins, self.true_destinations = None, None


    def create_matrix(self):
        """Creates distance matrix of points in meters"""

        query = self.gmaps.distance_matrix(self.true_origins, self.true_destinations, departure_time=self.now, mode="driving")
        for row in query['rows']:
            origin_to_dests = [el['distance']['value'] for el in row['elements']]
            self.distances.append(origin_to_dests)


    def get_distances(self, origins, destinations):
        """Retrieves current distance matrix between

        :param origins:
        :param destinations:

        """

        self.true_origins, self.true_destinations = origins, destinations
        self.create_matrix()


    def to_txt_matrix(self, has_headers=True):
        """Writes the Distance Matrix to .txt

        :param has_headers: add origins and destinations to the file, default is True

        """

        with open("dmatrix.txt", "w") as f:
            if has_headers:

                f.write('---'.join(d for d in self.true_destinations))
                f.write('\n')
                for o, row in zip(self.true_origins, self.distances):
                    f.write(o + '---')
                    for e in row:
                        f.write(str(e)+' ')
                    f.write('\n')
            else:

                for o, row in zip(self.true_origins, self.distances):
                    for e in row:
                        f.write(str(e)+' ')
                    f.write('\n')


    @staticmethod
    def combine_gmaps_address(row):
        """Returns GMaps address to search"""
        return row['Indirizzo'] + ', ' + row['Comune'] + ', IT'


    def read_from_excel(self, from_, to_):
        """Creates distance Matrix between

        :param from: excel file of origins
        :param to: excel file of destinations

        """

        from_df = pd.read_excel(from_)
        to_df = pd.read_excel(to_)

        from_df['gmaps'] = from_df.apply(lambda row: self.combine_gmaps_address(row), axis=1)
        to_df['gmaps'] = to_df.apply(lambda row: self.combine_gmaps_address(row), axis=1)

        self.true_origins, self.true_destinations = from_df['gmaps'].tolist(), to_df['gmaps'].tolist()
        self.create_matrix()



prova = DM()
prova.read_from_excel('origins.xlsx', 'dests.xlsx')
prova.to_txt_matrix(has_headers=False)
