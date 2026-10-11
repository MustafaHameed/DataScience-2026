<?php
/**
 * Zone Rate -- a shipping method written from nothing.
 *
 * Chapter 19. The flat rate that ships with the platform charges one amount
 * everywhere, which Chapter 11 showed is the method that is wrong about every
 * unusually distant order. This one charges per zone, with a free-delivery
 * threshold, which is what an actual Pakistani shop needs.
 *
 * THE WHOLE INTERFACE IS ONE METHOD. getQuote() is handed the delivery
 * address and returns either an empty array -- meaning "I do not serve this
 * address" -- or a structure describing one or more options. There is no base
 * class to extend and no interface to declare; the platform calls the method
 * if it exists. That is the entire extension point for shipping, and reading
 * flat.php next to this file is the fastest way to understand it.
 *
 * Two things it is easy to get wrong, both of which Chapter 11 measured:
 *
 *   * RETURN THE COST WITHOUT TAX, and give the tax class. The totals
 *     pipeline applies the tax, in its own position, and a method that
 *     returns a tax-inclusive cost is taxed twice.
 *
 *   * RESPECT THE GEO ZONE. If the shop has restricted this method to a geo
 *     zone and the address is outside it, return []. Returning a quote
 *     anyway is how a shop ends up promising delivery where it has no
 *     courier.
 */
namespace Opencart\Catalog\Model\Extension\Ecadmod\Shipping;

class ZoneRate extends \Opencart\System\Engine\Model {
	/**
	 * @param array<string, mixed> $address the delivery address
	 *
	 * @return array<string, mixed> one method with one or more quotes, or []
	 */
	public function getQuote(array $address): array {
		$this->load->language('extension/ecadmod/shipping/zonerate');

		// 1. Is this address inside the geo zone this method is limited to?
		//    A geo zone of 0 means "everywhere" -- see Chapter 11.
		$geo_zone_id = (int)$this->config->get('shipping_zonerate_geo_zone_id');

		if ($geo_zone_id) {
			$this->load->model('localisation/geo_zone');

			$in_zone = $this->model_localisation_geo_zone->getGeoZone(
				$geo_zone_id, (int)$address['country_id'], (int)$address['zone_id']);

			if (!$in_zone) {
				return [];
			}
		}

		// 2. Find the rate for this zone. The setting is a small text block,
		//    one "zone_id:cost" per line, because a shop with four provinces
		//    does not need a table and a user interface to maintain four
		//    numbers. Line "0:" is the fallback.
		$rates = [];

		foreach (explode("\n", (string)$this->config->get('shipping_zonerate_rates')) as $line) {
			$parts = explode(':', trim($line));

			if (count($parts) == 2 && $parts[0] !== '') {
				$rates[(int)$parts[0]] = (float)$parts[1];
			}
		}

		$zone_id = (int)$address['zone_id'];

		if (isset($rates[$zone_id])) {
			$cost = $rates[$zone_id];
		} elseif (isset($rates[0])) {
			$cost = $rates[0];
		} else {
			// No rate configured and no fallback: say nothing rather than
			// guessing. A missing quote is visible; a guessed one is not.
			return [];
		}

		// 3. Free over a threshold, measured on the sub-total BEFORE tax and
		//    before this charge -- which is the only basis a customer can
		//    check against the price list.
		$free_over = (float)$this->config->get('shipping_zonerate_free_over');
		$free = ($free_over > 0 && $this->cart->getSubTotal() >= $free_over);

		if ($free) {
			$cost = 0.0;
		}

		$tax_class_id = (int)$this->config->get('shipping_zonerate_tax_class_id');

		$quote_data = [];

		$quote_data['zonerate'] = [
			'code'         => 'zonerate.zonerate',
			'name'         => $free ? $this->language->get('text_free')
			                        : $this->language->get('text_description'),
			// Cost WITHOUT tax. The totals pipeline adds it.
			'cost'         => $cost,
			'tax_class_id' => $tax_class_id,
			'text'         => $this->currency->format(
				$this->tax->calculate($cost, $tax_class_id, $this->config->get('config_tax')),
				$this->session->data['currency'])
		];

		return [
			'code'       => 'zonerate',
			'name'       => $this->language->get('heading_title'),
			'quote'      => $quote_data,
			'sort_order' => $this->config->get('shipping_zonerate_sort_order'),
			'error'      => false
		];
	}
}
