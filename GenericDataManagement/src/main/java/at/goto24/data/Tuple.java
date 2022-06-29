package at.goto24.data;

/**
 * Created by p.taucher on 21.10.2016.
 */
public class Tuple<X, Y> {
    X x;
    Y y;

    public Tuple(X x, Y y) {
        this.x = x;
        this.y = y;
    }

    @Override
    public String toString() {
        return "(" + x + "," + y + ")";
    }

    @Override
    public boolean equals(Object other) {
        if (other == this) {
            return true;
        }

        if (!(other instanceof Tuple)) {
            return false;
        }

        Tuple<X, Y> other1 = (Tuple<X, Y>) other;

        boolean xEquals = false; 
        if (other1.x != null && this.x != null) {
            xEquals = other1.x.equals(this.x);
        } else if (other1.x == null && this.x == null) {
            xEquals = true;
        }

        boolean yEquals = false; 
        if (other1.y != null && this.y != null) {
            yEquals = other1.y.equals(this.y);
        } else if (other1.y == null && this.y == null) {
            yEquals = true;
        }
        
        return xEquals && yEquals;
    }

    @Override
    public int hashCode() {
        final int prime = 31;
        int result = 1;
        result = prime * result + ((x == null) ? 0 : x.hashCode());
        result = prime * result + ((y == null) ? 0 : y.hashCode());
        return result;
    }

    /**
     * @return x
     **/
    public X getX() {
        return x;
    }

    /**
     * @return y
     **/
    public Y getY() {
        return y;
    }

    /**
     * @param x
     */
    public void setX(X x) {
        this.x = x;
    }

    /**
     * @param y
     */
    public void setY(Y y) {
        this.y = y;
    }
}
