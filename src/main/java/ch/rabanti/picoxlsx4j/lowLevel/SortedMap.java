/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2019
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.picoxlsx4j.lowLevel;

import java.util.ArrayList;
import java.util.HashMap;

/**
 * Class representing a reduced sorted map (key / value). This class is not compatible with the Map interface
 * @author Raphael Stoeckli
 */
class SortedMap {

// ### P R I V A T E  F I E L D S ###    
    private int count;
    private final HashMap<String, Integer> index;
    private final ArrayList<String> keyEntries;
    private final ArrayList<String> valueEntries;
    
   
// ### C O N S T R U C T O R S ###
    /**
     * Default constructor
     */
    public SortedMap()
    {
        this.keyEntries = new ArrayList<>();
        this.valueEntries = new ArrayList<>();
        this.index = new HashMap<>();
        this.count = 0;
    }
    
// ### M E T H O D S ###    
    /**
     * Method to add a key value pair
     * @param key key as string
     * @param value value as string
     * @return returns the index of the inserted or replaced entry in the map
     */
    public int add(String key, String value)
    {
        if (this.index.containsKey(key))
        {
            return this.index.get(key);
        }
        else
        {
            this.index.put(key, this.count);
            this.keyEntries.add(key);
            this.valueEntries.add(value);
            this.count++;
            return this.count - 1;
        }
    }
    
    /**
     * Gets whether the specified key exists in the map
     * @param key Key to check
     * @return True if the entry exists, otherwise false
     */
    public boolean containsKey(String key)
    {
        return this.index.containsKey(key);
    }
    /**
     * Gets the value of the specified key
     * @param key Key of the entry
     * @return The value of the entry. If the key was not found, null is returned
     */
    public String get(String key)
    {
        if (this.index.containsKey(key))
        {
            return this.valueEntries.get(this.index.get(key));
        }
        return null;
    }
    /**
     * Gets the keys of the map as list
     * @return ArrayList of Keys
     */
    public ArrayList<String> getKeys()
    {
        return this.keyEntries;
    }
    /**
     * Gets the values of the map as list
     * @return ArrayList of Values
     */
    public ArrayList<String> getValues()
    {
        return this.valueEntries;
    }
    /**
     * Gets the size of the map
     * @return Number of entries in the map
     */
    public int size()
    {
        return this.count;
    }
    
}
