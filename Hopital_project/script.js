    let map;

function initMap() {
    map = new google.maps.Map(document.getElementById('map'), {
        center: {lat: -34.397, lng: 150.644},
        zoom: 8
    });
    getUserLocation();
}

function getUserLocation() {
    if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(function(position) {
            var userLocation = {
                lat: position.coords.latitude,
                lng: position.coords.longitude
            };
            map.setCenter(userLocation);
            var userMarker = new google.maps.Marker({
                position: userLocation,
                map: map,
                title: 'You are here',
                icon: 'https://maps.google.com/mapfiles/ms/icons/blue-dot.png'
        });
        }, function() {
            handleLocationError(true, map.getCenter());
        });
    } else {
        handleLocationError(false, map.getCenter());
    }
}

function handleLocationError(browserHasGeolocation, pos) {
    var infoWindow = new google.maps.InfoWindow;
    infoWindow.setPosition(pos);
    infoWindow.setContent(browserHasGeolocation ?
        'Error: The Geolocation service failed.' :
        'Error: Your browser doesn\'t support geolocation.');
    infoWindow.open(map);
}

function searchHospitals() {
    const location = document.getElementById('location').value;
    const disease = document.getElementById('disease').value;
    const day = document.getElementById('day').value;

    if (!location) {
        alert('Please enter a location.');
        return;
    }

    const service = new google.maps.places.PlacesService(map);
    service.textSearch({
        query: `${location} hospital ${disease}`,
        type: 'hospital'
    }, (results, status) => {
        if (status === google.maps.places.PlacesServiceStatus.OK) {
            results.forEach(place => {
                getPlaceDetails(place.place_id, day);
            });
            map.setCenter(results[0].geometry.location);
        } else {
            alert('No hospitals found. Please try a different location.');
        }
    });
}

function getPlaceDetails(placeId, day) {
    const service = new google.maps.places.PlacesService(map);
    service.getDetails({
        placeId: placeId,
        fields: ['name', 'formatted_address', 'geometry', 'formatted_phone_number']
    }, (place, status) => {
        if (status === google.maps.places.PlacesServiceStatus.OK) {
            createMarker(place, day);
        }
    });
}

function createMarker(place, day) {
    const marker = new google.maps.Marker({
        map: map,
        position: place.geometry.location
    });

    google.maps.event.addListener(marker, 'click', () => {
        const infoWindow = new google.maps.InfoWindow({
            content: `<div>
                        <strong>${place.name}</strong><br>
                        ${place.formatted_address}<br>
                        ${place.formatted_phone_number || ''}<br>
                        Available on: ${day}
                      </div>`
        });
        infoWindow.open(map, marker);
    });
}

window.onload = initMap;
