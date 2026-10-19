"""Sensor platform for the Outlook M365 integration."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from homeassistant.components.sensor import (
    SensorEntity,
    SensorEntityDescription,
    SensorStateClass,
)
from homeassistant.config_entries import ConfigEntry
from homeassistant.core import HomeAssistant
from homeassistant.helpers.device_registry import DeviceInfo
from homeassistant.helpers.entity_platform import AddEntitiesCallback
from homeassistant.helpers.update_coordinator import CoordinatorEntity

from .const import DATA_COORDINATOR, DOMAIN
from .coordinator import OutlookDataUpdateCoordinator


@dataclass(frozen=True, kw_only=True)
class OutlookSensorDescription(SensorEntityDescription):
    """Describes an Outlook sensor."""

    value_fn: Callable[[dict[str, Any]], Any]
    attributes_fn: Callable[[dict[str, Any]], dict[str, Any]] | None = None


def _latest_subject(data: dict[str, Any]) -> Any:
    latest = data.get("latest_message")
    return latest.get("subject") if latest else None


def _latest_attributes(data: dict[str, Any]) -> dict[str, Any]:
    latest = data.get("latest_message")
    if not latest:
        return {}
    sender = (latest.get("from") or {}).get("emailAddress", {})
    return {
        "message_id": latest.get("id"),
        "from_name": sender.get("name"),
        "from_address": sender.get("address"),
        "received": latest.get("receivedDateTime"),
        "is_read": latest.get("isRead"),
        "has_attachments": latest.get("hasAttachments"),
        "importance": latest.get("importance"),
        "preview": latest.get("bodyPreview"),
        "web_link": latest.get("webLink"),
    }


SENSORS: tuple[OutlookSensorDescription, ...] = (
    OutlookSensorDescription(
        key="unread_count",
        translation_key="unread_count",
        icon="mdi:email-alert",
        native_unit_of_measurement="messages",
        state_class=SensorStateClass.MEASUREMENT,
        value_fn=lambda data: data.get("unread_count"),
    ),
    OutlookSensorDescription(
        key="latest_subject",
        translation_key="latest_subject",
        icon="mdi:email",
        value_fn=_latest_subject,
        attributes_fn=_latest_attributes,
    ),
)


async def async_setup_entry(
    hass: HomeAssistant,
    entry: ConfigEntry,
    async_add_entities: AddEntitiesCallback,
) -> None:
    """Set up Outlook M365 sensors from a config entry."""
    coordinator: OutlookDataUpdateCoordinator = hass.data[DOMAIN][entry.entry_id][
        DATA_COORDINATOR
    ]
    async_add_entities(
        OutlookSensor(coordinator, entry, description) for description in SENSORS
    )


class OutlookSensor(CoordinatorEntity[OutlookDataUpdateCoordinator], SensorEntity):
    """Representation of an Outlook M365 sensor."""

    _attr_has_entity_name = True
    entity_description: OutlookSensorDescription

    def __init__(
        self,
        coordinator: OutlookDataUpdateCoordinator,
        entry: ConfigEntry,
        description: OutlookSensorDescription,
    ) -> None:
        """Initialise the sensor."""
        super().__init__(coordinator)
        self.entity_description = description
        self._attr_unique_id = f"{entry.entry_id}_{description.key}"
        self._attr_device_info = DeviceInfo(
            identifiers={(DOMAIN, entry.entry_id)},
            name="Outlook M365",
            manufacturer="Microsoft",
            model="Outlook / Microsoft 365",
            entry_type=None,
        )

    @property
    def native_value(self) -> Any:
        """Return the sensor value."""
        return self.entity_description.value_fn(self.coordinator.data or {})

    @property
    def extra_state_attributes(self) -> dict[str, Any] | None:
        """Return extra state attributes, if any."""
        if self.entity_description.attributes_fn is None:
            return None
        return self.entity_description.attributes_fn(self.coordinator.data or {})
