import numpy as np
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import math

try:
    from ortools.constraint_solver import routing_enums_pb2
    from ortools.constraint_solver import pywrapcp
except ImportError:
    pass

def get_retry_session(retries=8, backoff_factor=1.0):
    session = requests.Session()
    retry = Retry(total=retries, read=retries, connect=retries, backoff_factor=backoff_factor, status_forcelist=(400, 429, 500, 502, 503, 504))
    adapter = HTTPAdapter(max_retries=retry, pool_connections=100, pool_maxsize=100)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    return session

http_session = get_retry_session()

def calcular_matriz_distancias_numpy(coords):
    R = 6371000.0
    lats = np.radians(coords[:, 0])
    lons = np.radians(coords[:, 1])
    dlat = lats[:, np.newaxis] - lats
    dlon = lons[:, np.newaxis] - lons
    a = np.sin(dlat / 2.0)**2 + np.cos(lats)[:, np.newaxis] * np.cos(lats) * np.sin(dlon / 2.0)**2
    c = 2 * np.arcsin(np.sqrt(a))
    return (R * c).astype(int)

def obter_matriz_osrm(coords, url_osrm_base):
    if len(coords) > 100 and 'project-osrm' in url_osrm_base: return None
    coords_str = ";".join([f"{lon:.6f},{lat:.6f}" for lat, lon in coords])
    radiuses_str = ";".join(["10000"] * len(coords))
    url = f"{url_osrm_base}/table/v1/driving/{coords_str}?annotations=distance&radiuses={radiuses_str}"
    try:
        r = http_session.get(url, timeout=10)
        if r.status_code == 200:
            data = r.json()
            if data.get('code') == 'Ok': return np.array(data['distances']).astype(int).tolist()
    except: pass
    return None

def resolver_tsp_ortools(lista_obras, base_lat, base_lon, url_osrm_base):
    if not lista_obras: return []
    coords_array = np.array([(base_lat, base_lon)] + [(r['LATITUDE'], r['LONGITUDE']) for r in lista_obras])
    
    distance_matrix = obter_matriz_osrm(coords_array, url_osrm_base)
    if distance_matrix is None:
        distance_matrix = calcular_matriz_distancias_numpy(coords_array).tolist()
        
    manager = pywrapcp.RoutingIndexManager(len(distance_matrix), 1, 0)
    routing = pywrapcp.RoutingModel(manager)

    def distance_callback(from_index, to_index): return distance_matrix[manager.IndexToNode(from_index)][manager.IndexToNode(to_index)]
    transit_callback_index = routing.RegisterTransitCallback(distance_callback)
    routing.SetArcCostEvaluatorOfAllVehicles(transit_callback_index)
    
    penalty = 100000000 
    for node in range(1, len(distance_matrix)): routing.AddDisjunction([manager.NodeToIndex(node)], penalty)

    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    search_parameters.first_solution_strategy = routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC
    search_parameters.local_search_metaheuristic = routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH
    search_parameters.time_limit.seconds = 2 
    
    solution = routing.SolveWithParameters(search_parameters)
    rota_atual = []
    if solution:
        index = routing.Start(0)
        while not routing.IsEnd(index):
            node_index = manager.IndexToNode(index)
            if node_index != 0: rota_atual.append(lista_obras[node_index - 1])
            index = solution.Value(routing.NextVar(index))
    return rota_atual

def obter_rota_ruas(lat1, lon1, lat2, lon2, url_osrm_base, vel_fallback_kmh=30):
    if lat1 == lat2 and lon1 == lon2: return [[lon1, lat1], [lon2, lat2]], 0.0
    try:
        url = f"{url_osrm_base}/route/v1/driving/{lon1:.6f},{lat1:.6f};{lon2:.6f},{lat2:.6f}?overview=full&geometries=geojson&radiuses=10000;10000"
        r = http_session.get(url, timeout=20) 
        if r.status_code == 200 and r.json().get('code') == 'Ok':
            return r.json()['routes'][0]['geometry']['coordinates'], r.json()['routes'][0]['duration']
    except: pass
    coords = np.array([[lat1, lon1], [lat2, lon2]])
    dist_m = calcular_matriz_distancias_numpy(coords)[0][1]
    return [[lon1, lat1], [lon2, lat2]], (dist_m / 1000.0 / vel_fallback_kmh) * 3600
