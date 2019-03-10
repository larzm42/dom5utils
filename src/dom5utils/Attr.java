package dom5utils;

public class Attr implements Comparable<Attr> {
	private String key;
	private String value;
	
	public Attr(String key, String value) {
		super();
		this.key = key;
		this.value = value;
	}
	public Attr(String key, int value) {
		super();
		this.key = key;
		this.value = Integer.toString(value);
	}
	public String getKey() {
		return key;
	}
	public void setKey(String key) {
		this.key = key;
	}
	public String getValue() {
		return value;
	}
	public void setValue(String value) {
		this.value = value;
	}
	@Override
	public int compareTo(Attr o) {
		return key.compareTo(o.getKey());
	}
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((key == null) ? 0 : key.hashCode());
		return result;
	}
	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		Attr other = (Attr) obj;
		if (key == null) {
			if (other.key != null)
				return false;
		} else if (!key.equals(other.key))
			return false;
		return true;
	}
	
}
