
package fi.tuni.prog3.sisu;

import com.fasterxml.jackson.databind.JsonNode;

/**
 * Interface for extracting data from the Sisu API.
 */
public interface iAPI {
    /**
     * Returns a JsonObject that is extracted from the Sisu API.
     * @param urlString URL for retrieving information from the Sisu API.
     * @return JsonObject.
     */
    public JsonNode getJsonObjectFromApi(String urlString);
}
