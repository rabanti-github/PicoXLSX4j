/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2016
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j.lowLevel;

import java.util.ArrayList;

/**
 * Class representing a reduced sorted map (key / value). This class is not compatible with the Map interface
 * @author Raphael Stoeckli
 */
public class SortedMap {
    
    private ArrayList<Tuple> entries;
    
    /**
     * Default constructor
     */
    public SortedMap()
    {
        this.entries = new ArrayList<>();
    }
    
    /**
     * Method to add a key value pair
     * @param key key as string
     * @param value value as string
     * @return returns the index of the inserted or replaced entry in the map
     */
    public int add(String key, String value)
    {
        int s = this.entries.size();
        Tuple t;
        for(int i = 0; i < s; i++ )
        {
            t = this.entries.get(i);
            if (t.Key == key)
            {
                this.entries.set(i, new Tuple(key, value));
                return i;
            }
        }
        this.entries.add(new Tuple(key, value));
        return s;
    }
    
    /**
     * Gets the sze of the map
     * @return Number of entries in the map
     */
    public int size()
    {
        return this.entries.size();
    }
    
    /**
     * Sets / updates the entry at the specified index
     * @param index Index of the item
     * @param value New value
     * @return True if the element exists, otherwise false (was not updated)
     */
    public boolean set(int index, String value)
    {
        if (index < 0 || index >= this.entries.size()) { return false; }
        this.entries.set(index, new Tuple(this.entries.get(index).Key, value));
        return true;
    }
    
    /**
     * Gets the value of the specified key
     * @param key Key of the entry
     * @return The value of the entry. If the key was not found, null is returned
     */
    public String get(String key)
    {
        int s = this.entries.size();
        Tuple t;
        for(int i = 0; i < s; i++ )
        {
            t = this.entries.get(i);
            if (t.Key == key)
            {
                return this.entries.get(i).Value;
            }
        }
        return null;
    }
    
    /**
     * Gets whether the specified key exists in the map
     * @param key Key to check
     * @return True if the entry exists, otherwise false
     */
    public boolean containsKey(String key)
    {
        int s = this.entries.size();
        Tuple t;
        for(int i = 0; i < s; i++ )
        {
            t = this.entries.get(i);
            if (t.Key == key)
            {
                return true;
            }
        }
        return false;        
    }
    
    public void clear()
    {
        this.entries.clear();
    }
    
    /**
     * Gets an ArrayList of all keys. The keys are returned in the order of the addition of entries 
     * @return List of keys
     */
    public ArrayList<String> getKeys()
    {
        ArrayList<String> output = new ArrayList<>();
        int s = this.entries.size();
        for(int i = 0; i < s; i++ )
        {
            output.add(this.entries.get(i).Key);
        }
        return output;
    }
    
    /**
     * Gets an ArrayList of all values. The values are returned in the order of the addition of entries 
     * @return List of values
     */    
    public ArrayList<String> getValues()
    {
        ArrayList<String> output = new ArrayList<>();
        int s = this.entries.size();
        for(int i = 0; i < s; i++ )
        {
            output.add(this.entries.get(i).Value);
        }
        return output;
    }    
    
    /**
     * Sub-Class representing a tuple of key and value
     */
    public final static class Tuple
    {
        /**
         * Key of the tuple
         */
        public String Key;
        /**
         * Value of tuple
         */
        public String Value;
        
        /**
         * Default constor with parameters
         * @param key Key of the tuple
         * @param value Value of the tuple
         */
        public Tuple(String key, String value)
        {
            this.Key = key;
            this.Value = value;
        }
    }
    
    
}
