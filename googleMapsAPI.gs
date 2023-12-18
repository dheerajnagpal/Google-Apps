/*
 *
 * Google Maps Formulas for Google Sheets
 *
 * Written by Dheeraj Nagpal
 *
 * 
 *
 */

const md5 = (key = "") => {
  const code = key.toLowerCase().replace(/\s/g, "");
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key)
    .map((char) => (char + 256).toString(16).slice(-2))
    .join("");
};

const getCache = (key) => {
  return CacheService.getDocumentCache().get(md5(key));
};

const setCache = (key, value) => {
  const expirationInSeconds = 6 * 60 * 60; // max is 6 hours
  CacheService.getDocumentCache().put(md5(key), value, expirationInSeconds);
};

/**
 * Calculate the travel time between two locations
 * on Google Maps.
 *
 * =GOOGLEMAPS_DURATION("NY 10005", "Hoboken NJ", "walking")
 *
 * @param {String} origin The address of starting point
 * @param {String} destination The address of destination
 * @param {String} mode The mode of travel (driving, walking, bicycling or transit)
 * @return {String} The time in minutes
 * @customFunction
 */
const GOOGLEMAPS_DURATION = (origin, destination, mode = "driving") => {
  if (!origin || !destination) {
    throw new Error("No address specified!");
  }
  if (origin.map) {
    return origin.map(DISTANCE);
  }
  const key = ["duration", origin, destination, mode].join(",");
  const value = getCache(key);
  if (value !== null) return value;
  const { routes: [data] = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();
  if (!data) {
    throw new Error("No route found!");
  }
  const { legs: [{ duration: { text: time } } = {}] = [] } = data;
  setCache(key, time);
  return time;
};

/**
 * Calculate the distance between two
 * locations on Google Maps.
 *
 * =GOOGLEMAPS_DISTANCE("NY 10005", "Hoboken NJ", "walking")
 *
 * @param {String} origin The address of starting point
 * @param {String} destination The address of destination
 * @param {String} mode The mode of travel (driving, walking, bicycling or transit)
 * @return {String} The distance in miles
 * @customFunction
 */
const GOOGLEMAPS_DISTANCE = (origin, destination, mode = "driving") => {
  if (!origin || !destination) {
    throw new Error("No address specified!");
  }
  if (origin.map) {
    return origin.map(DISTANCE);
  }
  const key = ["distance", origin, destination, mode].join(",");
  const value = getCache(key);
  if (value !== null) return value;

  const { routes: [data] = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();
  if (!data) {
    throw new Error("No route found!");
  }
  const { legs: [{ distance: { text: distance } } = {}] = [] } = data;
  setCache(key, distance);

  return distance;
};

/**
 * Get the latitude and longitude of any
 * address on Google Maps.
 *
 * =GOOGLEMAPS_LATLONG("10 Hanover Square, NY")
 *
 * @param {String} address The address to lookup.
 * @return {String} The latitude and longitude of the address.
 * @customFunction
 */
const GOOGLEMAPS_LATLONG = (address) => {
  if (!address) {
    throw new Error("No address specified!");
  }
  if (address.map) {
    return address.map(LATLONG);
  }
  const key = ["latlong", address].join(",");
  const value = getCache(key);
  if (value !== null) return value;

  const { results: [data = null] = [] } = Maps.newGeocoder().geocode(address);
  if (data === null) {
    throw new Error("Address not found!");
  }
  const { geometry: { location: { lat, lng } } = {} } = data;
  const answer = `${lat}, ${lng}`;
  setCache(key, answer);
  return answer;
};

/**
 * Get the full address of any zip code or
 * partial address on Google Maps.
 *
 * =GOOGLEMAPS_ADDRESS("10005")
 *
 * @param {String} address The zip code or partial address to lookup.
 * @return {String} The full address from Google Maps
 * @customFunction
 */
const GOOGLEMAPS_ADDRESS = (address) => {
  if (!address) {
    throw new Error("No address specified!");
  }
  if (address.map) {
    return address.map(LATLONG);
  }
  const key = ["address", address].join(",");
  const value = getCache(key);
  if (value !== null) return value;

  const { results: [data = null] = [] } = Maps.newGeocoder().geocode(address);
  if (data === null) {
    throw new Error("Address not found!");
  }
  const { formatted_address } = data;
  setCache(key, formatted_address);
  return formatted_address;
};

/**
 * Use Reverse Geocoding to get the address of
 * a point location (latitude, longitude) on Google Maps.
 *
 * =GOOGLEMAPS_REVERSEGEOCODE(latitude, longitude)
 *
 * @param {String} latitude The latitude to lookup.
 * @param {String} longitude The longitude to lookup.
 * @return {String} The postal address of the point.
 * @customFunction
 */

const GOOGLEMAPS_REVERSEGEOCODE = (latitude, longitude) => {
  if (!latitude) {
    throw new Error("No latitude specified!");
  }
  if (!longitude) {
    throw new Error("No longitude specified!");
  }
  const key = ["reverse", latitude, longitude].join(",");
  const value = getCache(key);
  if (value !== null) return value;

  const { results: [data = {}] = [] } = Maps.newGeocoder().reverseGeocode(
    latitude,
    longitude
  );
  const { formatted_address } = data;
  setCache(key, formatted_address);
  return formatted_address;
};

/**
 * Get the country name of an address on Google Maps.
 *
 * =GOOGLEMAPS_COUNTRY("10 Hanover Square, NY")
 *
 * @param {String} address The address to lookup.
 * @return {String} The country of the address.
 * @customFunction
 */
const GOOGLEMAPS_COUNTRY = (address) => {
  if (!address) {
    throw new Error("No address specified!");
  }
  if (address.map) {
    return address.map(COUNTRY);
  }

  const key = ["country", address].join(",");
  const value = getCache(key);
  if (value !== null) return value;

  const { results: [data = null] = [] } = Maps.newGeocoder().geocode(address);
  if (data === null) {
    throw new Error("Address not found!");
  }
  const [{ short_name, long_name } = {}] = data.address_components.filter(
    ({ types: [level] }) => {
      return level === "country";
    }
  );
  if (!short_name) {
    throw new Error("Country not found!");
  }
  const answer = `${long_name} (${short_name})`;
  setCache(key, answer);
  return answer;
};

/**
 * Find the driving direction between two
 * locations on Google Maps.
 *
 * =GOOGLEMAPS_DIRECTIONS("NY 10005", "Hoboken NJ", "walking")
 *
 * @param {String} origin The address of starting point
 * @param {String} destination The address of destination
 * @param {String} mode The mode of travel (driving, walking, bicycling or transit)
 * @return {String} The driving direction
 * @customFunction
 */
const GOOGLEMAPS_DIRECTIONS = (origin, destination, mode = "driving") => {
  if (!origin || !destination) {
    throw new Error("No address specified!");
  }
  const key = ["directions", origin, destination, mode].join(",");
  const value = getCache(key);
  if (value !== null) return value;

  const { routes = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();
  if (!routes.length) {
    throw new Error("No route found!");
  }
  const directions = routes
    .map(({ legs }) => {
      return legs.map(({ steps }) => {
        return steps.map((step) => {
          return step.html_instructions
            .replace("><", "> <")
            .replace(/<[^>]+>/g, "");
        });
      });
    })
    .join(", ");
  setCache(key, directions);
  return directions;
};
