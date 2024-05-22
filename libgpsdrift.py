import requests
from pgeocode import haversine_distance

WIND_SOURCE_URL="https://windsaloft.us/winds.php"

naut_to_feet = lambda naut: naut*1.68781
feet_to_miles = lambda feet : feet*1.15078

class GPSDrift():
    drogue_descent_rate = None

    def __init__(self, launchsite: tuple[float], aero_profile: tuple[float], main_descent: float,\
            drogue_descent: float = None):
        """
        Initializer for GPSDrift
        ----
        launchsite: tuple(latitude, longitude, altitude)
        aero_profile: tuple(cross-sectional aera perpendicular to wind, coefficient of drag)
        main_descent: negative number representing the expected descent speed on main chute
        drogue_descent: (optional) do not set for single deployment; negative number represetning
                        the expected descent speed on drogue chute
        """
        self.set_launch_site(*launchsite)
        self.set_aerodynamic_profile(*aero_profile)
        self.set_main_descent_rate(main_descent)
        if not drogue_descent is None:
            self.set_drogue_descent_rate(drogue_descent)

    def set_launch_site(self, latitude: float, longitude: float, altitude: float):
        self.launchsite = (latitude, longitude, altitude)

    def set_aerodynamic_profile(self, cross_section_area: float, drag_coefficient: float):
        # TODO Update assert with proper error checking
        assert cross_section_area > 0 and drag_coefficient > 0
        self.cross_secction_area = cross_section_area
        self.drag_coefficient = drag_coefficient

    def set_main_descent_rate(self, descent_speed: float):
        # TODO Update assert with proper error checking
        assert descent_speed < 0
        self.main_descent_rate = descent_speed

    def set_drogue_descent_rate(self, descent_speed: float):
        # TODO Update assert with proper error checking
        assert descent_speed < 0
        self.drogue_descent_rate = descent_speed

    @staticmethod
    def get_winds(latitude: float, longitude: float, num_hour_offsets: int) -> list[dict]:
        """
        Retrieve data for windspeed from remote source
        """
        wind_list = []
        for i in range(0, num_hour_offsets+1):
            params = {"lat":latitude, "lon":longitude, "hourOffset":i}
            response = requests.get(WIND_SOURCE_URL, params=params)
            wind_list.append(params.update({"response":response}))
        return wind_list

    def linear_interpolation(self, x: float, x0: float, y0: float, x1: float, y1: float) -> float:
        if x1 - x0 > 0:
            return y0 + (x - x0) * ((y1 - y0) / (x1 - x0))
        else:
            raise ZeroDivisionError("x1 and x0 are equal, linear interpolation failed")

    def drift_distance(self, start_alt: float, end_alt: float, windspeed: float, descent_rate: float) -> float:
        """
        Calculate the lateral drift of the rocket on descent
        """
        traveled_altitude = start_alt - end_alt
        descent_time = traveled_altitude / descent_rate
        return descent_time * naut_to_feet(windspeed)

    def naut_distance_between(self, start_lat: float, start_lon: float, end_lat: float, end_lon: float) -> list[float]:
        return haversine_distance((start_lat, start_lon), (end_lat), (end_lon))

