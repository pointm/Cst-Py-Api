def CstPickFace(mws, Name, componentname, id):
    pick = mws.Pick

    pick.PickFaceFromId((componentname + ":" + Name), str(id))
