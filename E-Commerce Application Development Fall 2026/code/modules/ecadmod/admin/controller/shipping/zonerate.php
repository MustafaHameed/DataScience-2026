<?php
/**
 * Zone Rate -- the admin half.
 *
 * Chapter 19. A shipping method's settings are shop-wide, so unlike the
 * module they live in oc_setting and are read back through $this->config.
 * That is the difference between the two kinds of extension in this folder,
 * and it is worth being deliberate about which one you are writing: a module
 * is placed, possibly several times, each with its own settings; a shipping
 * method is configured once.
 */
namespace Opencart\Admin\Controller\Extension\Ecadmod\Shipping;

class ZoneRate extends \Opencart\System\Engine\Controller {
	public function index(): void {
		$this->load->language('extension/ecadmod/shipping/zonerate');

		$this->document->setTitle($this->language->get('heading_title'));

		$data['breadcrumbs'] = [];
		$data['breadcrumbs'][] = [
			'text' => $this->language->get('text_home'),
			'href' => $this->url->link('common/dashboard', 'user_token=' . $this->session->data['user_token'])
		];

		$data['save'] = $this->url->link('extension/ecadmod/shipping/zonerate.save', 'user_token=' . $this->session->data['user_token']);
		$data['back'] = $this->url->link('marketplace/extension', 'user_token=' . $this->session->data['user_token'] . '&type=shipping');

		$data['shipping_zonerate_rates'] = $this->config->get('shipping_zonerate_rates');
		$data['shipping_zonerate_free_over'] = $this->config->get('shipping_zonerate_free_over');
		$data['shipping_zonerate_tax_class_id'] = $this->config->get('shipping_zonerate_tax_class_id');
		$data['shipping_zonerate_geo_zone_id'] = $this->config->get('shipping_zonerate_geo_zone_id');
		$data['shipping_zonerate_status'] = $this->config->get('shipping_zonerate_status');
		$data['shipping_zonerate_sort_order'] = $this->config->get('shipping_zonerate_sort_order');

		$this->load->model('localisation/tax_class');
		$data['tax_classes'] = $this->model_localisation_tax_class->getTaxClasses();

		$this->load->model('localisation/geo_zone');
		$data['geo_zones'] = $this->model_localisation_geo_zone->getGeoZones();

		$data['header'] = $this->load->controller('common/header');
		$data['column_left'] = $this->load->controller('common/column_left');
		$data['footer'] = $this->load->controller('common/footer');

		$this->response->setOutput($this->load->view('extension/ecadmod/shipping/zonerate', $data));
	}

	public function save(): void {
		$this->load->language('extension/ecadmod/shipping/zonerate');

		$json = [];

		if (!$this->user->hasPermission('modify', 'extension/ecadmod/shipping/zonerate')) {
			$json['error']['warning'] = $this->language->get('error_permission');
		}

		if (!$json) {
			$this->load->model('setting/setting');

			$this->model_setting_setting->editSetting('shipping_zonerate', $this->request->post);

			$json['success'] = $this->language->get('text_success');
		}

		$this->response->addHeader('Content-Type: application/json');
		$this->response->setOutput(json_encode($json));
	}
}
