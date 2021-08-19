
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import models.EntityData;
import parsers.JsonParser;

import java.io.File;
import java.io.IOException;
import java.util.*;

public class CodeAnalyzer {

    public static void main(String args[]) throws IOException {

        List<EntityData> entityDataList= JsonParser.parse("sample.json");
        ExcelGenerator.createExcelFile("",entityDataList);

    }


}
