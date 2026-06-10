"use client";

import React from "react";

export default function EquipmentMusterScreen(props) {
  const {
    styles,
    activeEquipmentDepartmentLabel,
    equipmentDepartment,
    EquipmentMusterModule,
    getEquipmentDisplayImage,
    getEquipmentFallbackImage,
    getImageUrl,
    getShipDisplayName,
    isAdmin,
    loadMasterInventoryItems,
    makeInventoryShip,
    masterInventoryLoading,
    musterItems,
    musterMessage,
    pictureLibraryBusy,
    pictureLibraryMessage,
    setEquipmentMode,
    syncMasterInventoryPicturesFromDrive,
    uploadEquipmentPictureZipFile,
    uploadMusterFile,
    userShip,
  } = props;

  return (
      <EquipmentMusterModule
        styles={styles}
        userShip={userShip}
        makeInventoryShip={makeInventoryShip}
        isAdmin={isAdmin}
        equipmentDepartment={equipmentDepartment}
        activeEquipmentDepartmentLabel={activeEquipmentDepartmentLabel}
        getShipDisplayName={getShipDisplayName}
        musterItems={musterItems}
        musterMessage={musterMessage}
        pictureLibraryMessage={pictureLibraryMessage}
        pictureLibraryBusy={pictureLibraryBusy}
        masterInventoryLoading={masterInventoryLoading}
        uploadMusterFile={uploadMusterFile}
        loadMasterInventoryItems={loadMasterInventoryItems}
        syncMasterInventoryPicturesFromDrive={syncMasterInventoryPicturesFromDrive}
        uploadEquipmentPictureZipFile={uploadEquipmentPictureZipFile}
        getEquipmentDisplayImage={getEquipmentDisplayImage}
        getEquipmentFallbackImage={getEquipmentFallbackImage}
        getImageUrl={getImageUrl}
        onBack={() => setEquipmentMode("")}
      />
    );
}
