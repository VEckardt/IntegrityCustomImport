package com.ptc.services.gateway.importer.excel;

import com.mks.gateway.mapper.ItemMapperException;
import com.mks.gateway.mapper.config.ItemMapperConfig;
import com.ptc.services.common.gateway.MappingConfig;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 *
 * @author Craig Stoss
 *
 */
public class Content {

    /**
     * The mapping of Header Value to Content Value
     */
    private HashMap<String, String> fieldValuesMap = new HashMap<String, String>();
    /**
     * If position is specified the String representing the position
     */
    private String position;

    private int rowNum;
    /**
     * Any Child contents.
     */
    private List<Content> children = new ArrayList<Content>();
    /**
     * Any Child contents.
     */
    private List<Content> relationships = new ArrayList<Content>();

    public String getExternalID(ItemMapperConfig mappingConfig) throws ItemMapperException {
        for (String key : fieldValuesMap.keySet()) {
            if (key.contentEquals("External ID") || MappingConfig.getInternalFieldName (mappingConfig, key).contentEquals("External ID")) {
                return fieldValuesMap.get(key);
            }
        }
        return "";
    }

    /**
     * Constructs a Content Object which has a position and has zero or more
     * Field Name to Field Value values.
     *
     * @param position the "position" as found in the position column of the
     * spreadsheet
     * @param fieldValuesMap the map of Field Names to Field values
     * @param rowNum
     */
    public Content(String position, HashMap<String, String> fieldValuesMap, int rowNum) {
        this.position = position;
        this.fieldValuesMap = fieldValuesMap;
        this.rowNum = rowNum;
    }

    /**
     * Constructs a Content Object which has no position and has zero or more
     * Field Name to Field Value values. This should be used when creating a
     * flat structure document.
     *
     * @param fieldValuesMap the map of Field Names to Field values
     * @param rowNum
     */
    public Content(HashMap<String, String> fieldValuesMap, int rowNum) {
        this.position = null;
        this.fieldValuesMap = fieldValuesMap;
        this.rowNum = rowNum;
    }

    /**
     * The children of this Content in order of addition.
     *
     * @return the list of all of the Children Content or an empty list if there
     * are none.
     */
    public List<Content> getChildren() {
        return children;
    }

    /**
     * This returns an unmodifiable Set of the children. By definition this will
     * give you the unique children as defined by {@link #equals(Object)}. This
     * is simply a convenience method should you wish to remove duplication.
     *
     * @return an unmodifiable set of the children.
     */
    public final Set<Content> getUniqueChildren() {
        return Collections.unmodifiableSet(new HashSet<Content>(children));
    }

    /**
     * Returns a sorted List of children. The sort order is determined by the
     * Comparator <tt>comparator</tt>. This is a convenience method to allow you
     * to provide some ordering to the Content items. (for example based off one
     * of the field values in {@link #fieldValuesMap}.
     * <br><b>NOTE</b>: This permanently sorts the children List. By calling
     * this method you are re-ordering {@link #children}. This is irreversible.
     *
     * @param comparator the comparator to use to sort the children.
     * @return the sorted list of children
     */
    public final List<Content> getChildren(Comparator<Content> comparator) {
        Collections.sort(children, comparator);
        return children;
    }

    /**
     * Set the children of this object.
     *
     * @param children the List of children to set for this parent.
     */
    public void setChildren(List<Content> children) {
        this.children = children;
    }

    /**
     * Retrieve the Field Values map. This is the map of the Field names to the
     * Field values that will be turned into Field Elements in the IIF.
     *
     * @return the map of Field names to Field values
     */
    public final HashMap<String, String> getFieldValuesMap() {
        return fieldValuesMap;
    }

    /**
     * Retrieve the "position" of this Content.
     *
     * @return the position of this Content
     */
    public final String getPosition() {
        return position;
    }

    public int getRowId() {
        return rowNum;
    }

    public int getRowNum() {
        return rowNum + 1;
    }

    /**
     * getSectionLevel, assuming the section is like 1, 1.1, 1.2, 2, 2.1, 2.2
     * etc.
     *
     * @return
     */
    public final int getSectionLevel() {
        try {
            String pos = position;
            if (pos.endsWith(".0")) {
                pos = pos.substring(0, pos.length() - 2);
            }
            return pos.replaceAll("[^.]", "").length() + 1;
        } catch (NullPointerException ex) {
            return 1;
        }
    }

    /**
     * @return the relationships
     */
    public List<Content> getRelationships() {
        return relationships;
    }

    /**
     * @param relationships the relationships to set
     */
    public void setRelationships(List<Content> relationships) {
        this.relationships = relationships;
    }

    public boolean addRelationship(Content content) {
        return relationships.add(content);
    }

    public boolean hasRelationships() {
        return relationships != null ? (relationships.size() > 0) : false;
    }

    /**
     * Add a new child Content to this parent.
     *
     * @param content the child to add {@link #fieldValuesMap}
     * @return as per {@link java.util.List#add(Object)}
     */
    public boolean addChild(Content content) {
        return children.add(content);
    }

    /**
     * Add a child at a specific index.
     *
     * @param index the index of the child
     * @param content the child to add
     */
    public void addChild(int index, Content content) {
        children.add(index, content);
    }

    /**
     * A convenient way to view the parent's position and children.
     *
     * @return
     */
    public String toString() {
        return position + "[" + children + "]";
    }
    /* (non-Javadoc)
     * @see java.lang.Object#hashCode()
     */

    @Override
    public int hashCode() {
        final int prime = 31;
        int result = 1;
        result = prime * result
                + ((fieldValuesMap == null) ? 0 : fieldValuesMap.hashCode());
        result = prime * result
                + ((position == null) ? 0 : position.hashCode());
        return result;
    }

    /**
     * Determine equality of this Content Object with another. Content that is
     * equal is defined as any content where the position and all of the Field
     * Name to Field Value mappings are equivalent.<br>
     * ie. {@link #fieldValuesMap} and {@link #position} are equivalent using
     * the respective {@link java.util.HashMap#equals(Object)} and
     * {@link java.lang.String#equals(Object)} methods.
     *
     * @param obj the Object to check for equality
     * @return
     */
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null) {
            return false;
        }
        if (getClass() != obj.getClass()) {
            return false;
        }
        Content other = (Content) obj;
        if (fieldValuesMap == null) {
            if (other.fieldValuesMap != null) {
                return false;
            }
        } else if (!fieldValuesMap.equals(other.fieldValuesMap)) {
            return false;
        }
        if (position == null) {
            if (other.position != null) {
                return false;
            }
        } else if (!position.equals(other.position)) {
            return false;
        }
        return true;
    }

}
