package parsers;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import models.EntityData;
import models.FieldData;

import java.io.File;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.atomic.AtomicReference;

public class JsonParser {

    private static void traverseAllNodes(JsonNode jsonNode, String currentNode, List<EntityData> entityDataList) {
        if (!jsonNode.isValueNode()) {
            Iterator<Map.Entry<String, JsonNode>> fieldIterator = jsonNode.fields();
            Iterator<JsonNode> elementsIterator = jsonNode.elements();
            List<String> nodesLst = new ArrayList<>();
            String fieldHeaders = "";
            String fieldValues = "";
            EntityData newEntity = new EntityData();
            while (fieldIterator.hasNext()) {
                Map.Entry<String, JsonNode> fields = fieldIterator.next();
                if (!fields.getValue().isValueNode()) {
                    //   System.out.println(currentNode+ " node " + fields.getKey());
                    String node = fields.getKey();
                    System.out.println("???"+fields.getValue());
                    //entityDataList.add(newEntity);
                    nodesLst.add(node);
                } else {
                     FieldData fieldData = new FieldData();
                    fieldData.setFieldName(fields.getValue().asText());
                    fieldData.setFieldHeaderName(fields.getKey());
                    newEntity.getFieldsList().add(fieldData);
                    fieldHeaders += fields.getKey() + ",";
                    fieldValues += fields.getValue() + ",";

                    // System.out.println(currentNode +" field " + fields.getKey() + fields.getValue());
                }
            }
            if (currentNode != null) {
                newEntity.setEntityName(findCurrentNodeName(currentNode));
                String parentNodeName = findParentNodeName(currentNode);
                Optional<EntityData> entityDataOptional = findEntity(entityDataList, parentNodeName);
                int depth=currentNode.split("->").length-1;
                newEntity.setDepth(depth);
                if (entityDataOptional.isPresent()) {
                    boolean isNodePresent = checkNodeListHasEntity(entityDataOptional.get(), newEntity.getEntityName());
                    if (!entityDataOptional.get().getEntityName().equals(newEntity.getEntityName()) && (!isNodePresent)) {
                        entityDataOptional.get().getEntityDataList().add(newEntity);
                    }
                    if (isNodePresent && !fieldHeaders.isEmpty()) {
                        entityDataOptional.get().getEntityDataList().add(newEntity);
                    }
                } else {
                    entityDataList.add(newEntity);
                }
                System.out.println(currentNode + "\n field Keys " + fieldHeaders + "\n field values " + fieldValues);
            }

            int index = 0;
            while (elementsIterator.hasNext()) {
                JsonNode nextNode = elementsIterator.next();

                if (!nextNode.isValueNode()) {
                    String currNode = "";
                    if (currentNode != null) {
                        currNode = currentNode;
                    }
                    if (index < nodesLst.size()) {
                        currNode += "->" + nodesLst.get(index);
                    }
                    traverseAllNodes(nextNode, currNode, entityDataList);
                    index++;
                }
            }

        } else {
            System.out.println("## " + jsonNode.asText());
        }
    }

    private static boolean checkNodeListHasEntity(EntityData entityData, String entityName) {
        return entityData.getEntityDataList().stream().filter(entity -> entity.getEntityName().equals(entityName)).findFirst().isPresent();
    }

    private static Optional<EntityData> findEntity(List<EntityData> entityDataList, String entityName) {

        AtomicReference<Optional<EntityData>> entityDataOptional = new AtomicReference<>(entityDataList.stream().filter(entityData -> {
            return entityData.getEntityName().equals(entityName);
        }).findFirst());
        if (!entityDataOptional.get().isPresent()) {
            entityDataList.forEach(entityData -> {
                        entityDataOptional.set(findEntity(entityData.getEntityDataList(), entityName));

                    }
            );
        }
        return entityDataOptional.get();
    }

    private static String findParentNodeName(String entityName) {
        String splitStrings[] = entityName.split("->");
        if (splitStrings.length > 2) {
            return splitStrings[splitStrings.length - 2];
        } else {
            return splitStrings[1];
        }
    }

    private static String findCurrentNodeName(String entityName) {
        String splitStrings[] = entityName.split("->");
        if (splitStrings.length > 2) {
            return splitStrings[splitStrings.length - 1];
        } else {
            return splitStrings[1];
        }
    }

    public static void main(String args[]) throws IOException {
        System.out.println(findCurrentNodeName("->customer"));
        System.out.println(findCurrentNodeName("->customer->shipTo"));
        System.out.println(findCurrentNodeName("->customer->shipTo->addresses->address"));
        parse("sample.json");

    }

    public static List<EntityData> parse(String fileName) throws IOException {
        ObjectMapper objectMapper = new ObjectMapper();
        File jsonInputFile = new File(fileName);
        JsonNode jsonNode = objectMapper.readTree(jsonInputFile);
        List<EntityData> entityDataList = new ArrayList<>();
        traverseAllNodes(jsonNode, null, entityDataList);
        return entityDataList;
    }
}
