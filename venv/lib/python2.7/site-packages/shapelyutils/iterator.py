

def iterate(shapely_obj, skip_empties=True):
    """Allows iterating over pretty much anything only returning Point, LineString and Polygon objects.
    Handy for simplifying your work with shapely.

    Note that empty elements, even though they python object is a Polygon, will report
    geom_type to be GeometryCollection!

    Keyword arguments:
    skip_empties -- True by default. Whether or not to skip empty elements, which are basically non-iterable GeometryCollections.
    """
    if shapely_obj is not None:
        if hasattr(shapely_obj, "__iter__"):
            for sub_obj in shapely_obj:
                yield from iterate(sub_obj, skip_empties=skip_empties)
        elif hasattr(shapely_obj, "geom_type"):
            if not shapely_obj.is_empty:
                if shapely_obj.geom_type in ("Point", "LineString", "Polygon"):
                    yield shapely_obj
                else:
                    raise RuntimeError("Unexpected geometry type %s"%shapely_obj.geom_type)
            elif not skip_empties:
                yield shapely_obj


