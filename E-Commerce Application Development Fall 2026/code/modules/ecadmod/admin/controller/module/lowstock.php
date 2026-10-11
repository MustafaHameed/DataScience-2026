<?php
/**
 * Low Stock -- the admin half.
 *
 * Chapter 19. A storefront module needs two controllers: one that renders it
 * for a customer, and one that lets the shopkeeper configure it. They share a
 * name and nothing else.
 *
 * This one is longer than the storefront half, which is the usual proportion
 * and worth expecting: most of the work in extending a platform is the
 * administration, not the feature.
 *
 * The pattern below is OpenCart's and is worth following exactly:
 *
 *   index()  builds the form, including the breadcrumbs and the save URL.
 *   save()   validates, writes, answers with JSON. It never renders.
 *
 * Note that save() checks a permission before it does anything. Forgetting
 * that line is the single commonest security defect in third-party
 * extensions for this platform -- the form is only shown to an administrator,
 * so the author assumes the action is too, and the action is a URL.
 */
namespace Opencart\Admin\Controller\Extension\Ecadmod\Module;

class LowStock extends \Opencart\System\Engine\Controller {
	public function index(): void {
		$this->load->language('extension/ecadmod/module/lowstock');

		$this->document->setTitle($this->language->get('heading_title'));

		$module_id = isset($this->request->get['module_id'])
			? (int)$this->request->get['module_id'] : 0;

		$data['breadcrumbs'] = [];
		$data['breadcrumbs'][] = [
			'text' => $this->language->get('text_home'),
			'href' => $this->url->link('common/dashboard', 'user_token=' . $this->session->data['user_token'])
		];
		$data['breadcrumbs'][] = [
			'text' => $this->language->get('heading_title'),
			'href' => $this->url->link('extension/ecadmod/module/lowstock', 'user_token=' . $this->session->data['user_token'] . '&module_id=' . $module_id)
		];

		$data['save'] = $this->url->link('extension/ecadmod/module/lowstock.save', 'user_token=' . $this->session->data['user_token'] . '&module_id=' . $module_id);
		$data['back'] = $this->url->link('marketplace/extension', 'user_token=' . $this->session->data['user_token'] . '&type=module');

		// Defaults, then whatever was saved for THIS instance. The module can
		// be placed more than once with different settings, so nothing here
		// may come from $this->config.
		$setting = ['name' => '', 'threshold' => 5, 'limit' => 4, 'status' => 0];

		if ($module_id) {
			$this->load->model('setting/module');
			$info = $this->model_setting_module->getModule($module_id);

			if ($info) {
				$setting = array_merge($setting, $info);
			}
		}

		$data['module_id']  = $module_id;
		$data['name']       = $setting['name'];
		$data['threshold']  = $setting['threshold'];
		$data['limit']      = $setting['limit'];
		$data['status']     = $setting['status'];

		$data['header'] = $this->load->controller('common/header');
		$data['column_left'] = $this->load->controller('common/column_left');
		$data['footer'] = $this->load->controller('common/footer');

		$this->response->setOutput($this->load->view('extension/ecadmod/module/lowstock', $data));
	}

	public function save(): void {
		$this->load->language('extension/ecadmod/module/lowstock');

		$json = [];

		// Check this FIRST. The form is only shown to someone with
		// permission; the action is a URL anyone can post to.
		if (!$this->user->hasPermission('modify', 'extension/ecadmod/module/lowstock')) {
			$json['error']['warning'] = $this->language->get('error_permission');
		}

		$post = $this->request->post + ['name' => '', 'threshold' => 5, 'limit' => 4, 'status' => 0];

		if (oc_strlen($post['name']) < 3 || oc_strlen($post['name']) > 64) {
			$json['error']['name'] = $this->language->get('error_name');
		}

		if (!$json) {
			$this->load->model('setting/module');

			$data = [
				'name'      => (string)$post['name'],
				'threshold' => max(1, (int)$post['threshold']),
				'limit'     => max(1, min(20, (int)$post['limit'])),
				'status'    => (int)$post['status']
			];

			if (empty($this->request->get['module_id'])) {
				$this->model_setting_module->addModule('ecadmod.lowstock', $data);
			} else {
				$this->model_setting_module->editModule((int)$this->request->get['module_id'], $data);
			}

			$json['success'] = $this->language->get('text_success');
		}

		$this->response->addHeader('Content-Type: application/json');
		$this->response->setOutput(json_encode($json));
	}
}
